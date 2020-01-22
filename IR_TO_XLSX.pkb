create or replace PACKAGE BODY IR_TO_XLSX
as

  CURSOR cur_highlight(
    p_report_id in APEX_APPLICATION_PAGE_IR_RPT.REPORT_ID%TYPE,
    p_delimetered_column_list IN VARCHAR2
  ) 
  IS
  SELECT rez.* ,
       rownum cond_number,
       'HIGHLIGHT_'||rownum cond_name
  FROM (
  SELECT report_id,
         CASE 
           WHEN condition_operator in ('not in', 'in') THEN
             ir_to_xlsx.get_highlight_in_cond_sql(condition_expression,condition_sql,condition_column_name)
           ELSE 
             REPLACE(REPLACE(REPLACE(REPLACE(condition_sql,'#APXWS_EXPR#',''''||condition_expression||''''),'#APXWS_EXPR2#',''''||condition_expression2||''''),'#APXWS_HL_ID#','1'),'#APXWS_CC_EXPR#','"'||condition_column_name||'"') 
         END condition_sql,
         condition_column_name,
         condition_enabled,
         highlight_row_color,
         highlight_row_font_color,
         highlight_cell_color,
         highlight_cell_font_color      
    FROM apex_application_page_ir_cond
    WHERE condition_type = 'Highlight'
      AND report_id = p_report_id
      AND instr(':'||p_delimetered_column_list||':',':'||condition_column_name||':') > 0
      AND condition_enabled = 'Yes'
    ORDER BY --rows highlights first 
             nvl2(highlight_row_color,1,0) DESC, 
             nvl2(highlight_row_font_color,1,0) DESC,
             highlight_sequence 
    ) rez;

  SUBTYPE t_color IS VARCHAR2(7);
  SUBTYPE t_style_string IS VARCHAR2(300);
  SUBTYPE t_format_mask IS VARCHAR2(100);
  SUBTYPE t_font IS VARCHAR2(50);
  SUBTYPE t_large_varchar2 IS VARCHAR2(32767);
  SUBTYPE largevarchar2 IS VARCHAR2(32767);
  SUBTYPE columntype IS VARCHAR2(15);
  SUBTYPE t_formatmask IS VARCHAR2(100);

  TYPE t_date_table IS TABLE OF DATE INDEX BY BINARY_INTEGER;  
  TYPE t_number_table IS TABLE OF NUMBER INDEX BY BINARY_INTEGER;  
  TYPE t_styles IS TABLE OF BINARY_INTEGER INDEX BY t_style_string;
  TYPE t_fill_list IS TABLE OF BINARY_INTEGER INDEX BY t_color;
  TYPE t_font_list IS TABLE OF BINARY_INTEGER INDEX BY t_font;
  TYPE t_format_mask_list IS TABLE OF BINARY_INTEGER INDEX BY t_format_mask;

  TYPE t_col_names IS TABLE OF apex_application_page_ir_col.report_label%TYPE INDEX BY apex_application_page_ir_col.column_alias%TYPE;
  TYPE t_col_format_mask IS TABLE OF apex_application_page_ir_comp.computation_format_mask%TYPE INDEX BY apex_application_page_ir_col.column_alias%TYPE;
  TYPE t_header_alignment IS TABLE OF apex_application_page_ir_col.heading_alignment%TYPE INDEX BY apex_application_page_ir_col.column_alias%TYPE;
  TYPE t_column_alignment IS TABLE OF apex_application_page_ir_col.column_alignment%TYPE INDEX BY apex_application_page_ir_col.column_alias%TYPE;
  TYPE t_column_data_types IS TABLE OF apex_application_page_ir_col.column_type%TYPE INDEX BY BINARY_INTEGER;
  TYPE t_highlight IS TABLE OF cur_highlight%rowtype INDEX BY BINARY_INTEGER;
  TYPE t_column_links IS TABLE OF apex_application_page_ir_col.column_link%TYPE INDEX BY apex_application_page_ir_col.column_alias%TYPE;
  TYPE t_all_columns IS TABLE OF BINARY_INTEGER INDEX BY VARCHAR2(32767);

  TYPE ir_report IS RECORD
   (
    report                    apex_ir.t_report,
    ir_data                   APEX_APPLICATION_PAGE_IR_RPT%ROWTYPE,
    displayed_columns         APEX_APPLICATION_GLOBAL.VC_ARR2,
    break_on                  APEX_APPLICATION_GLOBAL.VC_ARR2,
    break_really_on           APEX_APPLICATION_GLOBAL.VC_ARR2, -- "break on" except hidden columns
    sum_columns_on_break      APEX_APPLICATION_GLOBAL.VC_ARR2,
    avg_columns_on_break      APEX_APPLICATION_GLOBAL.VC_ARR2,
    max_columns_on_break      APEX_APPLICATION_GLOBAL.VC_ARR2,
    min_columns_on_break      APEX_APPLICATION_GLOBAL.VC_ARR2,
    median_columns_on_break   APEX_APPLICATION_GLOBAL.VC_ARR2,
    count_columns_on_break    APEX_APPLICATION_GLOBAL.VC_ARR2,
    count_distnt_col_on_break APEX_APPLICATION_GLOBAL.VC_ARR2,
    skipped_columns           BINARY_INTEGER default 0, -- when scpecial columns like apxws_row_pk is used
    start_with                BINARY_INTEGER default 0, -- position of first displayed column in query
    end_with                  BINARY_INTEGER default 0, -- position of last displayed column in query    
    agg_cols_cnt              BINARY_INTEGER default 0, 
    column_header_labels      t_col_names,       -- column names in report header
    col_format_mask           t_col_format_mask, -- format like $3849,56
    row_highlight             t_highlight,
    col_highlight             t_highlight,
    header_alignment          t_header_alignment,
    column_alignment          t_column_alignment,
    column_data_types         t_column_data_types,
    column_link               t_column_links,
    desc_tab                  DBMS_SQL.DESC_TAB2, -- information about columns from final sql query
    all_columns               t_all_columns      -- list of all columns from final sql query     
   );  

   TYPE t_cell_data IS RECORD
   (
     VALUE           VARCHAR2(100),
     text            largevarchar2,
     datatype        VARCHAR2(50)
   );  

  -- EXCEPTIONS
  format_error EXCEPTION;
  PRAGMA EXCEPTION_INIT(format_error, -01830);
  date_format_error EXCEPTION;
  PRAGMA EXCEPTION_INIT(date_format_error, -01821);
  conversion_error EXCEPTION;
  PRAGMA EXCEPTION_INIT(conversion_error,-06502);

  l_report           ir_report;   

  -- these plSQL tables are used ONLY to identify that unique entity
  -- was already used,not to keep the XML content.
  g_styles           t_styles;  
  g_fills            t_fill_list;  
  g_fonts            t_font_list;  
  g_number_formats   t_format_mask_list;
  -- the unique XML content is saved in these text variables 
  g_fonts_xml        t_large_varchar2;
  g_back_xml         t_large_varchar2;
  v_styles_xml       CLOB;
  v_format_mask_xml  CLOB;

  -- CONSTANT
  STRING_HEIGHT CONSTANT NUMBER DEFAULT 14.4;
  WIDTH_COEFFICIENT CONSTANT NUMBER DEFAULT 6;
  LOG_PREFIX VARCHAR2(9) DEFAULT 'GPV_XLSX:';
  BACK_COLOR CONSTANT t_color DEFAULT '#C6E0B4';  
  -- Excel-template should have some pre-defined
  -- fonts, format masks, fills, styles etc. to format column headers.
  -- They are hardcoded in constants below (BOLD_FONT,UNDERLINE_FONT etc.).
  -- One should keep this entities by adding new one. 
  -- Constatns below define an amount of this hardcoded entities 
  FONTS_CNT CONSTANT  BINARY_INTEGER DEFAULT 3;  
  FORMAT_MASK_CNT CONSTANT BINARY_INTEGER DEFAULT 1; 
  DEFAULT_FILL_CNT CONSTANT  BINARY_INTEGER DEFAULT 2; 
  DEFAULT_STYLES_CNT CONSTANT  BINARY_INTEGER DEFAULT 1;     
  FORMAT_MASK_START_WITH  CONSTANT  BINARY_INTEGER DEFAULT 164;  

  ------------------------------------------------------------------------------
  t_sheet_rels CLOB DEFAULT '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  </Relationships>';

  t_workbook CLOB DEFAULT '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4506"/>
    <workbookPr filterPrivacy="1" defaultThemeVersion="124226"/>
    <bookViews>
      <workbookView xWindow="120" yWindow="120" windowWidth="24780" windowHeight="12150"/>
    </bookViews>
    <sheets>
      <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
    <definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sheet1!$A$1:$H$1</definedName></definedNames>
    <calcPr calcId="125725"/>
    <fileRecoveryPr repairLoad="1"/>
  </workbook>';

  t_style_template CLOB DEFAULT '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
    #FORMAT_MASKS#
    #FONTS#
    #FILLS#
    <borders count="2">
      <border>
        <left/>
        <right/>
        <top/>
        <bottom/>
        <diagonal/>
      </border>
      <border>
        <left/>  
        <right/> 
          <top style="thin">
        <color indexed="64"/>
        </top>
        <bottom/>
        <diagonal/>
      </border>
    </borders>
    <cellStyleXfs count="2">
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
      <xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0" />
    </cellStyleXfs>
    #STYLES#
    <cellStyles count="2">
      <cellStyle name="Link" xfId="1" builtinId="8" />
      <cellStyle name="Standard" xfId="0" builtinId="0" />
    </cellStyles>    
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
    <extLst>
      <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
        <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
      </ext>
    </extLst>
  </styleSheet>';

  DEFAULT_FONT CONSTANT VARCHAR2(200) DEFAULT '
   <font>
      <sz val="11" />
      <color theme="1" />
      <name val="Calibri" />
      <family val="2" />
      <scheme val="minor" />
    </font>';

  BOLD_FONT CONSTANT VARCHAR2(200) DEFAULT '
   <font>
      <b />
      <sz val="11" />
      <color theme="1" />
      <name val="Calibri" />
      <family val="2" />
      <scheme val="minor" />
    </font>';
  UNDERLINE_FONT CONSTANT VARCHAR2(200) DEFAULT '
   <font>
      <u />
      <sz val="11" />
      <color theme="10" />
      <name val="Calibri" />
      <family val="2" />
      <scheme val="minor" />
    </font>';    

  DEFAULT_FILL CONSTANT VARCHAR2(200) DEFAULT  '
    <fill>
      <patternFill patternType="none" />
    </fill>
    <fill>
      <patternFill patternType="gray125" />
    </fill>'; 

  DEFAULT_STYLE CONSTANT VARCHAR2(200) DEFAULT '
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';

  AGGREGATE_STYLE CONSTANT VARCHAR2(250) DEFAULT '
      <xf numFmtId="#FMTID#" borderId="1" fillId="0" fontId="1" xfId="0" applyAlignment="1" applyFont="1" applyBorder="1">
         <alignment wrapText="1" horizontal="right"  vertical="top"/>
     </xf>';

  HEADER_STYLE CONSTANT VARCHAR2(200) DEFAULT '     
      <xf numFmtId="0" borderId="0" #FILL# fontId="1" xfId="0" applyAlignment="1" applyFont="1" >
         <alignment wrapText="0" horizontal="#ALIGN#"/>
     </xf>';    

  FORMAT_MASK CONSTANT VARCHAR2(100) DEFAULT '
      <numFmt numFmtId="'||FORMAT_MASK_START_WITH||'" formatCode="dd.mm.yyyy"/>';

  ------------------------------------------------------------------------------
  /**
  * http://mk-commi.blogspot.co.at/2014/11/concatenating-varchar2-values-into-clob.html  

  * Procedure concatenates a VARCHAR2 to a CLOB.
  * It uses another VARCHAR2 as a buffer until it reaches 32767 characters.
  * Then it flushes the current buffer to the CLOB and resets the buffer using
  * the actual VARCHAR2 to add.
  * Your final call needs to be done setting p_eof to TRUE in order to
  * flush everything to the CLOB.
  *
  * @param p_clob        The CLOB buffer.
  * @param p_vc_buffer   The intermediate VARCHAR2 buffer. (must be VARCHAR2(32767))
  * @param p_vc_addition The VARCHAR2 value you want to append.
  * @param p_eof         Indicates if complete buffer should be flushed to CLOB.
  */
  PROCEDURE add(
   p_clob IN OUT NOCOPY CLOB
  ,p_vc_buffer IN OUT NOCOPY VARCHAR2
  ,p_vc_addition IN VARCHAR2
  ,p_eof IN BOOLEAN DEFAULT false
  ) AS
  BEGIN
    -- Standard Flow
    IF nvl(lengthb(p_vc_buffer), 0) + nvl(lengthb(p_vc_addition), 0) < (32767/2) THEN
      p_vc_buffer := p_vc_buffer || convert(p_vc_addition,'utf8');
    ELSE
      IF p_clob IS NULL THEN
        dbms_lob.createtemporary(p_clob, true);
      END IF;
      dbms_lob.writeappend(p_clob, length(p_vc_buffer), p_vc_buffer);
      p_vc_buffer := convert(p_vc_addition,'utf8');
    END IF;

    -- Full Flush requested
    IF p_eof THEN
      IF p_clob IS NULL THEN
         p_clob := p_vc_buffer;
      ELSE
        dbms_lob.writeappend(p_clob, length(p_vc_buffer), p_vc_buffer);
      END IF;
      p_vc_buffer := NULL;
    END IF;
  END add;
  ------------------------------------------------------------------------------
  PROCEDURE log(p_message IN VARCHAR2)
  IS
  BEGIN
    apex_debug_message.log_message(p_message => log_prefix||substr(p_message,1,32758),
                                   p_enabled => false,
                                   p_level   => 4);
  END log; 
  ------------------------------------------------------------------------------
  -- convert date format string from ORACLE-style to MS Excel style
  FUNCTION convert_date_format(p_format IN VARCHAR2)
  RETURN VARCHAR2
  IS
    v_str      VARCHAR2(100);
  BEGIN
    v_str := apex_plugin_util.replace_substitutions(p_format);
    v_str := upper(v_str);
    --date
    v_str := replace(v_str,'DAY','DDDD');
    v_str := replace(v_str,'MONTH','MMMM');
    v_str := replace(v_str,'MON','MMM');
    v_str := replace(v_str,'R','Y');    
    v_str := replace(v_str,'FM','');
    v_str := replace(v_str,'PM',' AM/PM');
    --time
    v_str := replace(v_str,'MI','mm');
    v_str := replace(v_str,'SS','ss');

    v_str := replace(v_str,'HH24','hh');
    v_str := regexp_replace(v_str,'HH12([^ ]+)','h\1 AM/PM');    
    v_str := replace(v_str,'AM\/PM',' AM/PM');
    RETURN v_str;
  END convert_date_format;
  ------------------------------------------------------------------------------
  -- convert number format string from ORACLE-style to MS Excel style
  FUNCTION convert_number_format(p_format IN VARCHAR2)
  RETURN VARCHAR2
  IS
    v_str      VARCHAR2(100);
  BEGIN
    v_str := apex_plugin_util.replace_substitutions(p_format);
    v_str := upper(v_str);
    --number
    v_str := replace(v_str,'9','#');
    v_str := replace(v_str,'D','.');
    v_str := replace(v_str,'G',',');
    v_str := replace(v_str,'FM','');
    -- plus/minus sign
    IF instr(v_str,'S') > 0 THEN
      v_str := replace(v_str,'S','');
      v_str := '+'||v_str||';-'||v_str;
    END IF;

    --currency
    v_str := replace(v_str,'L',convert('&quot;'||rtrim(TO_CHAR(0,'FML0'),'0')||'&quot;','UTF8'));    

    v_str := regexp_substr(v_str,'.G?[^G]+$');
    RETURN v_str;
  END convert_number_format;  
  ------------------------------------------------------------------------------
  -- get current date/time format used in DB session for given date or time type
  -- this format is used if date/time format in APEX is not defined 
  FUNCTION get_current_format(p_data_type IN BINARY_INTEGER)
  RETURN VARCHAR2
  IS
    v_format    t_formatmask;
    v_parameter VARCHAR2(50);
  BEGIN
    IF p_data_type IN (dbms_types.typecode_timestamp_tz,181) THEN
       v_parameter := 'NLS_TIMESTAMP_TZ_FORMAT';
    ELSIF p_data_type IN (dbms_types.typecode_timestamp_ltz,231) THEN
       v_parameter := 'NLS_TIMESTAMP_TZ_FORMAT';
    ELSIF p_data_type IN (dbms_types.typecode_timestamp,180) THEN
       v_parameter := 'NLS_TIMESTAMP_FORMAT';
    ELSIF p_data_type = dbms_types.typecode_date THEN
       v_parameter := 'NLS_DATE_FORMAT';
    ELSE
       RETURN 'dd.mm.yyyy';
    END IF;

    SELECT value
    INTO v_format
    FROM v$nls_parameters
    WHERE parameter = v_parameter;

    RETURN v_format;
  END get_current_format;
  ------------------------------------------------------------------------------
  -- convert date format from ORACLE-style to MS Excel Style
  -- for given date/time datatype 
  FUNCTION convert_date_format(
    p_datatype IN VARCHAR2,
    p_format IN VARCHAR2
  )
  RETURN VARCHAR2
  IS
    v_str        VARCHAR2(100);
    v_24h_format  BOOLEAN DEFAULT true;
  BEGIN
    IF p_format IS NULL THEN
      IF p_datatype = 'TIMESTAMP_TZ' THEN
        v_str := get_current_format(dbms_types.typecode_timestamp_tz);
      ELSIF p_datatype = 'TIMESTAMP_LTZ' THEN
        v_str := get_current_format(dbms_types.typecode_timestamp_ltz);  
      ELSIF p_datatype = 'TIMESTAMP' THEN
        v_str := get_current_format(dbms_types.typecode_timestamp);          
      ELSIF p_datatype = 'DATE' THEN
       v_str := get_current_format(dbms_types.typecode_date);
      ELSE
        v_str := get_current_format('');
      END IF;
    ELSE
      v_str := p_format;
    END IF;
    RETURN convert_date_format(p_format => v_str);
  END convert_date_format;  
  ------------------------------------------------------------------------------
  -- convert date format from ORACLE-style to format used in JavaScript
  -- for given date/time datatype   
  FUNCTION convert_date_format_js(
    p_datatype IN VARCHAR2, 
    p_format IN VARCHAR2
   )
  RETURN VARCHAR2
  IS
    v_str        VARCHAR2(100);
    v_24h_format  BOOLEAN DEFAULT true;
  BEGIN
    IF p_format IS NULL THEN
      IF p_datatype = 'TIMESTAMP_TZ' THEN
        v_str := get_current_format(dbms_types.typecode_timestamp_tz);
      ELSIF p_datatype = 'TIMESTAMP_LTZ' THEN
        v_str := get_current_format(dbms_types.typecode_timestamp_ltz);  
      ELSIF p_datatype = 'TIMESTAMP' THEN
        v_str := get_current_format(dbms_types.typecode_timestamp);          
      ELSIF p_datatype = 'DATE' THEN
       v_str := get_current_format(dbms_types.typecode_date);
      ELSE
        v_str := get_current_format('');
      END IF;
    ELSE
      v_str := p_format;
    END IF;

    v_str := upper(apex_plugin_util.replace_substitutions(v_str));
    --date
    v_str := replace(v_str,'DAY','ddd');
    v_str := replace(v_str,'MONTH','MMMM');
    v_str := replace(v_str,'MON','MMM');
    v_str := replace(v_str,'R','Y');
    v_str := replace(v_str,'FM','');
    IF instr(v_str,'PM') > 0 THEN
      v_24h_format := false;
    END IF;
    v_str := replace(v_str,'AM/PM','');
    v_str := replace(v_str,'PM','');
    --time
    v_str := replace(v_str,'MI','mm');
    v_str := replace(v_str,'SS','ss');
    v_str := replace(v_str,'HH24','HH');
    v_str := replace(v_str,'HH12','hh');
    IF NOT v_24h_format THEN
      v_str := replace(v_str,'HH','hh');
    END IF;

    RETURN v_str;
  END convert_date_format_js;
  ------------------------------------------------------------------------------
  -- to set a font-color in Excel one need to define a Font 
  -- add_font function:
  -- return an index of the Font based on the combination (font-color + has link property)
  -- create a new Font if Font not exists and return an id
  FUNCTION add_font(
    p_font_color IN t_color,
    p_is_link    IN BOOLEAN 
  )
  RETURN BINARY_INTEGER 
  IS  
   v_index t_font;
  BEGIN
    IF p_is_link IS NOT NULL THEN
      v_index := p_font_color||'L';
    ELSIF p_font_color IS NOT NULL THEN
      v_index := p_font_color;
    ELSE 
      raise_application_error(-20001,'Font is null and Link is null!');
    END IF;

    IF NOT g_fonts.EXISTS(v_index) THEN
      g_fonts(v_index) := g_fonts.count + 1;
      g_fonts_xml := g_fonts_xml||
        '<font>'||chr(10)||
        CASE WHEN p_is_link IS NOT NULL THEN 
        '   <u />'||chr(10)
        ELSE ''||
        '   <sz val="11" />'||chr(10)
        END ||
        CASE WHEN p_font_color IS NOT NULL THEN 
        '   <color rgb="'||ltrim(p_font_color,'#')||'" />'||chr(10)
        ELSE -- link        
        '   <color theme="10" />'||chr(10)
        END ||
        '   <name val="Calibri" />'||chr(10)||
        '   <family val="2" />'||chr(10)||
        '   <scheme val="minor" />'||
        '</font>'||chr(10);

        RETURN g_fonts.count + FONTS_CNT - 1; --start with 0
     ELSE
       RETURN g_fonts(v_index) + FONTS_CNT - 1;
     END IF;
  END add_font;
  ------------------------------------------------------------------------------
  -- to set a background-color in Excel one need create a Fill 
  -- add_fill function:
  -- return an index of the Fill based on the background-color
  -- create a new Fill if Fill not exists and return an id  
  FUNCTION  add_fill(p_back_color IN t_color)
  RETURN BINARY_INTEGER 
  IS
  BEGIN
    IF NOT g_fills.EXISTS(p_back_color) THEN
      g_fills(p_back_color) := g_fills.count + 1;
      g_back_xml := g_back_xml||
        '<fill>'||chr(10)||
        '   <patternFill patternType="solid">'||chr(10)||
        '     <fgColor rgb="'||ltrim(p_back_color,'#')||'" />'||chr(10)||
        '     <bgColor indexed="64" />'||chr(10)||
        '   </patternFill>'||
        '</fill>'||chr(10);

        RETURN  g_fills.count + DEFAULT_FILL_CNT - 1; --start with 0
     ELSE
       RETURN g_fills(p_back_color) + DEFAULT_FILL_CNT - 1;
     END IF;
  END add_fill;
  ------------------------------------------------------------------------------  
  -- to set a Number-Format in Excel one need create them 
  -- add_number_format function:
  -- return an index of the Number-Format based on the format mask
  -- create a new Number-Format if Number-Format not exists and return an id  
  FUNCTION  add_number_format(p_mask IN VARCHAR2)
  RETURN BINARY_INTEGER 
  IS
  BEGIN
    IF NOT g_number_formats.EXISTS(p_mask) THEN
      g_number_formats(p_mask) := g_number_formats.count + 1;
      v_format_mask_xml := v_format_mask_xml||
        '<numFmt numFmtId="'||(FORMAT_MASK_START_WITH + g_number_formats.count)||'" formatCode="'||p_mask||'"/>'||chr(10);
        RETURN  g_number_formats.count + FORMAT_MASK_CNT + FORMAT_MASK_START_WITH - 1; 
     ELSE
       RETURN g_number_formats(p_mask) + FORMAT_MASK_CNT + FORMAT_MASK_START_WITH - 1;
     END IF;
  END add_number_format;  
  ------------------------------------------------------------------------------
  -- build a Fonts: XML-list of Font-elements
  FUNCTION get_fonts_xml 
  RETURN CLOB
  IS
  BEGIN
    RETURN to_clob('<fonts count="'||(g_fonts.count + FONTS_CNT)||'" x14ac:knownFonts="1">'||
                   DEFAULT_FONT||chr(10)||
                   BOLD_FONT||chr(10)||
                   UNDERLINE_FONT||chr(10)||
                   g_fonts_xml||chr(10)||
                   '</fonts>'||chr(10));
  END get_fonts_xml;  
  ------------------------------------------------------------------------------
  -- build a Fills: XML-list of Fill-elements
  FUNCTION get_fills_xml 
  RETURN CLOB
  IS
  BEGIN
    RETURN to_clob('<fills count="'||(g_fills.count + DEFAULT_FILL_CNT)||'">'||
                   default_fill||
                   g_back_xml||chr(10)||
                   '</fills>'||chr(10));
  END get_fills_xml;  
  ------------------------------------------------------------------------------ 
  -- build a numFmts: XML-list of numFmt-elements
  FUNCTION get_numFmts_xml 
  RETURN CLOB
  IS
  BEGIN
    RETURN to_clob('<numFmts count="'||(g_number_formats.count + FORMAT_MASK_CNT)||'">'||
                   FORMAT_MASK||
                   v_format_mask_xml||chr(10)||
                   '</numFmts>'||chr(10));
  END get_numFmts_xml;  
  ------------------------------------------------------------------------------  
  -- build a cellXfs: XML-list of cell styles
  FUNCTION get_cellxfs_xml
  RETURN CLOB
  IS
  BEGIN
    RETURN to_clob('<cellXfs count="'||(g_styles.count + default_styles_cnt)||'">'||
                   DEFAULT_STYLE||chr(10)||
                   v_styles_xml||chr(10)||
                   '</cellXfs>'||chr(10));
  END get_cellxfs_xml;
  ------------------------------------------------------------------------------ 
  -- get an index of existent Numer- or Date- Format
  -- create the new Format if not exist 
  FUNCTION get_num_fmt_id(p_data_type   IN VARCHAR2,
                          p_format_mask IN VARCHAR2)
  RETURN BINARY_INTEGER
  IS
    v_format_mask VARCHAR2(100);    
  BEGIN
      v_format_mask := apex_plugin_util.replace_substitutions(p_format_mask);
      IF p_data_type = 'NUMBER' THEN 
       IF v_format_mask IS NULL THEN
          RETURN 0; -- default number format in Excel
        ELSE
          RETURN add_number_format(convert_number_format(v_format_mask)); 
        END IF;
      ELSIF p_data_type = 'DATE' THEN 
        IF v_format_mask IS NULL THEN
          RETURN FORMAT_MASK_START_WITH;  -- default date format
        ELSE
          RETURN add_number_format(convert_date_format(v_format_mask)); 
        END IF;
      ELSE
        RETURN 2;  -- default string format in Excel
      END IF;  

  END get_num_fmt_id;  
  ------------------------------------------------------------------------------  
  -- get style_id for existent styles or 
  -- add new style and return style_id
  FUNCTION get_style_id(
    p_font_color  IN t_color,
    p_back_color  IN t_color,
    p_data_type   IN VARCHAR2,
    p_format_mask IN VARCHAR2,
    p_align       IN VARCHAR2,
    p_is_link     IN BOOLEAN
  )
  RETURN BINARY_INTEGER
  IS
    v_style           t_style_string;
    v_font_color_id   BINARY_INTEGER;
    v_back_color_id   BINARY_INTEGER;
    v_style_xml       t_large_varchar2;
  BEGIN
    -- generate unique style name based on cell format
    v_style := nvl(p_font_color,'      ')
       ||nvl(p_back_color,'       ')
       ||p_data_type
       ||p_format_mask
       ||p_align
       ||CASE WHEN p_is_link THEN 'LINK' ELSE '' END;

    IF g_styles.EXISTS(v_style) THEN
      RETURN g_styles(v_style)   - 1 + default_styles_cnt;
    ELSE
      g_styles(v_style) := g_styles.count + 1;

      IF p_is_link THEN
        v_style_xml := '<xf borderId="0" xfId="0" ';
      ELSE
        v_style_xml := '<xf borderId="0" xfId="1" ';
      END IF;
      v_style_xml := v_style_xml||replace(' numFmtId="#FMTID#" ','#FMTID#',get_num_fmt_id(p_data_type,p_format_mask));

      IF p_font_color IS NOT NULL OR p_is_link THEN
        v_font_color_id := add_font(p_font_color,p_is_link);
        v_style_xml := v_style_xml||' fontId="'||v_font_color_id||'"  applyFont="1" ';
      ELSE
        v_style_xml := v_style_xml||' fontId="0" '; --default font
      END IF;

      IF p_back_color IS NOT NULL THEN
        v_back_color_id := add_fill(p_back_color);
        v_style_xml := v_style_xml||' fillId="'||v_back_color_id||'"  applyFill="1" ';
      ELSE
        v_style_xml := v_style_xml||' fillId="0" '; --default fill 
      END IF;

      v_style_xml := v_style_xml||'>'||chr(10);
      v_style_xml := v_style_xml||'<alignment wrapText="1"';
      IF p_align IS NOT NULL THEN
        v_style_xml := v_style_xml||' horizontal="'||lower(p_align)||'" ';
      END IF;
      v_style_xml := v_style_xml||'/>'||chr(10);

      v_style_xml := v_style_xml||'</xf>'||chr(10);    
      v_styles_xml := v_styles_xml||to_clob(v_style_xml);     

      RETURN g_styles.count  - 1 + DEFAULT_STYLES_CNT;
    END IF;  

  END get_style_id;
  ------------------------------------------------------------------------------  
  -- analog of get_style_id for aggregate functions of Interactive Report
  FUNCTION get_aggregate_style_id(
    p_font        IN t_color,                               
    p_back        IN t_color,
    p_data_type   IN VARCHAR2,
    p_format_mask IN VARCHAR2
  )
  RETURN BINARY_INTEGER
  IS
    v_style           t_style_string;
    v_style_xml       t_large_varchar2;
    v_num_fmt_id      BINARY_INTEGER;
  BEGIN
    -- generate unique style name based on cell format
    v_style := nvl(p_font,'      ')||nvl(p_back,'       ')||p_data_type||p_format_mask||'AGGREGATE';
    --   
    IF g_styles.exists(v_style) THEN
      RETURN g_styles(v_style)  - 1 + DEFAULT_STYLES_CNT;
    ELSE
      g_styles(v_style) := g_styles.count + 1;
      v_style_xml  := replace(AGGREGATE_STYLE,'#FMTID#',get_num_fmt_id(p_data_type,p_format_mask)) ||chr(10);      
      v_styles_xml := v_styles_xml||v_style_xml;      
      RETURN g_styles.count  - 1 + DEFAULT_STYLES_CNT;
    END IF;  
  END get_aggregate_style_id;  
  ------------------------------------------------------------------------------
  -- analog of get_style_id for report headers
  FUNCTION get_header_style_id(
    p_back        IN t_color DEFAULT BACK_COLOR,
    p_align       IN VARCHAR2,
    p_border      IN BOOLEAN DEFAULT FALSE
  )
  RETURN BINARY_INTEGER
  IS
    v_style             t_style_string;
    v_style_xml         t_large_varchar2;
    v_num_fmt_id        BINARY_INTEGER;
    v_fill_id           BINARY_INTEGER;
  BEGIN
    -- generate unique style name based on cell format
    v_style := '      '||nvl(p_back,'       ')||'CHARHEADER'||p_align;
    --   
    IF g_styles.exists(v_style) THEN
      RETURN g_styles(v_style)  - 1 + DEFAULT_STYLES_CNT;
    ELSE
      g_styles(v_style) := g_styles.count + 1;

      IF p_back IS NOT NULL THEN
        v_fill_id := add_fill(p_back);
        v_style_xml := replace(HEADER_STYLE,'#FILL#',' fillId="'||v_fill_id||'"  applyFill="1" ' );
      ELSE
        v_style_xml := replace(HEADER_STYLE,'#FILL#',' fillId="0" '); --default fill 
      END IF;      
      v_style_xml  := replace(v_style_xml,'#ALIGN#',lower(p_align))||chr(10);
      v_styles_xml := v_styles_xml||v_style_xml;      
      RETURN g_styles.count  - 1 + DEFAULT_STYLES_CNT;
    END IF;  
  END get_header_style_id;    
  ------------------------------------------------------------------------------
  -- build a final XML containing all styles  
  FUNCTION  get_styles_xml
  RETURN CLOB
  IS
    v_template CLOB;
  BEGIN
     v_template := replace(t_style_template,'#FORMAT_MASKS#',get_numfmts_xml);     
     v_template := replace(v_template,'#FONTS#',get_fonts_xml);
     v_template := replace(v_template,'#FILLS#',get_fills_xml);
     v_template := replace(v_template,'#STYLES#',get_cellxfs_xml);

     RETURN v_template;
  END get_styles_xml;  
  ------------------------------------------------------------------------------
  -- get cell name in Excel Format (like A11 ob B5)
  FUNCTION get_cell_name(
    p_col IN BINARY_INTEGER,
    p_row IN BINARY_INTEGER
  )
  RETURN VARCHAR2
  IS   
  BEGIN
  /*
   Author: Moritz Klein (https://github.com/commi235)
   https://github.com/commi235/xlsx_builder/blob/master/xlsx_builder_pkg.pkb
  */
      RETURN CASE
                WHEN p_col > 702
                THEN
                      CHR (64 + TRUNC ( (p_col - 27) / 676))
                   || CHR (65 + MOD (TRUNC ( (p_col - 1) / 26) - 1, 26))
                   || CHR (65 + MOD (p_col - 1, 26))
                WHEN p_col > 26
                THEN
                   CHR (64 + TRUNC ( (p_col - 1) / 26)) || CHR (65 + MOD (p_col - 1, 26))
                ELSE
                   CHR (64 + p_col)
             END||p_row;
  END get_cell_name;
  ----------------------------------------------------------- 
  FUNCTION get_column_header_label(p_column_alias IN apex_application_page_ir_col.column_alias%TYPE)
  RETURN apex_application_page_ir_col.report_label%TYPE
  IS
  BEGIN
    -- https://github.com/glebovpavel/IR_to_MSExcel/issues/9
    -- Thanks HeavyS 
    RETURN  apex_plugin_util.replace_substitutions(p_value => l_report.column_header_labels(p_column_alias));
  EXCEPTION
    WHEN OTHERS THEN
       raise_application_error(-20001,'get_column_header_label: p_column_alias='||p_column_alias||' '||SQLERRM);
  END get_column_header_label;
  ------------------------------------------------------------------------------
  FUNCTION get_col_format_mask(p_column_alias IN apex_application_page_ir_col.column_alias%TYPE)
  RETURN t_formatmask
  IS
  BEGIN
    RETURN apex_plugin_util.replace_substitutions(
      p_value => replace(l_report.col_format_mask(p_column_alias),'"',''),
      p_escape => false);
  EXCEPTION
    WHEN OTHERS THEN
       raise_application_error(-20001,'get_col_format_mask: p_column_alias='||p_column_alias||' '||SQLERRM);
  END get_col_format_mask;
  ------------------------------------------------------------------------------
  FUNCTION get_col_link(p_column_alias IN apex_application_page_ir_col.column_alias%TYPE)
  RETURN t_formatmask
  IS
  BEGIN
    RETURN replace(l_report.column_link(p_column_alias),'"','');
  EXCEPTION
    WHEN OTHERS THEN
       raise_application_error(-20001,'get_col_link: p_column_alias='||p_column_alias||' '||SQLERRM);
  END get_col_link;
  ------------------------------------------------------------------------------
  PROCEDURE set_col_format_mask(
    p_column_alias IN apex_application_page_ir_col.column_alias%TYPE,
    p_format_mask  IN t_formatmask
  )
  IS
  BEGIN
      l_report.col_format_mask(p_column_alias) := p_format_mask;
  EXCEPTION
    WHEN OTHERS THEN
       raise_application_error(-20001,'set_col_format_mask: p_column_alias='||p_column_alias||' '||sqlerrm);
  END set_col_format_mask;
  ------------------------------------------------------------------------------
  FUNCTION get_header_alignment(p_column_alias IN apex_application_page_ir_col.column_alias%TYPE)
  RETURN apex_application_page_ir_col.heading_alignment%TYPE
  IS
  BEGIN
    RETURN l_report.header_alignment(p_column_alias);
  EXCEPTION
    WHEN OTHERS THEN
       raise_application_error(-20001,'get_header_alignment: p_column_alias='||p_column_alias||' '||sqlerrm);
  END get_header_alignment;
  ------------------------------------------------------------------------------
  FUNCTION get_column_alignment(p_column_alias IN apex_application_page_ir_col.column_alias%TYPE)
  RETURN apex_application_page_ir_col.column_alignment%TYPE
  IS
  BEGIN
    RETURN l_report.column_alignment(p_column_alias);
  EXCEPTION
    WHEN OTHERS THEN
       raise_application_error(-20001,'get_column_alignment: p_column_alias='||p_column_alias||' '||sqlerrm);
  END get_column_alignment;
  ------------------------------------------------------------------------------
  FUNCTION get_column_data_type(p_num IN BINARY_INTEGER)
  RETURN apex_application_page_ir_col.column_type%TYPE
  IS
  BEGIN
    RETURN l_report.column_data_types(p_num);
  EXCEPTION
    WHEN OTHERS THEN
       raise_application_error(-20001,'get_column_data_type: p_num='||p_num||' '||sqlerrm);
  END get_column_data_type;
  ------------------------------------------------------------------------------  
  FUNCTION get_column_alias(p_num IN BINARY_INTEGER)
  RETURN VARCHAR2
  IS
  BEGIN
    RETURN l_report.displayed_columns(p_num);
  EXCEPTION
    WHEN OTHERS THEN
       raise_application_error(-20001,'get_column_alias: p_num='||p_num||' '||sqlerrm);
  END get_column_alias;
  ------------------------------------------------------------------------------
  -- get column  alias by column number from SQL-Query 
  FUNCTION get_column_alias_sql(p_num IN BINARY_INTEGER) -- column number in sql-query                               
  RETURN VARCHAR2
  IS
  BEGIN
    RETURN l_report.displayed_columns(p_num - l_report.start_with + 1);
  EXCEPTION
    WHEN OTHERS THEN
       raise_application_error(-20001,'get_column_alias_sql: p_num='||p_num||' '||sqlerrm);
  END get_column_alias_sql;
  ------------------------------------------------------------------------------
  -- get value from varchar2-table by position 
  FUNCTION get_current_row(
    p_current_row    IN apex_application_global.vc_arr2,
    p_position       IN BINARY_INTEGER
  )
  RETURN largevarchar2
  IS
  BEGIN
    RETURN p_current_row(p_position);
  EXCEPTION
    WHEN OTHERS THEN
       raise_application_error(-20001,'get_current_row: string: p_id='||p_position||' '||sqlerrm);
  END get_current_row; 
  ------------------------------------------------------------------------------
  -- get value from date2-table by position 
  FUNCTION get_current_row(
    p_current_row IN t_date_table,
    p_position    IN BINARY_INTEGER
  )
  RETURN DATE
  IS
  BEGIN
    IF p_current_row.EXISTS(p_position) THEN
      RETURN p_current_row(p_position);
    ELSE
      RETURN NULL;
    END IF;
  EXCEPTION
    WHEN OTHERS THEN
       raise_application_error(-20001,'get_current_row:date: p_id='||p_position||' '||sqlerrm);
  END get_current_row;   
  ------------------------------------------------------------------------------
  -- get value from numbers-table by position 
  FUNCTION get_current_row(
    p_current_row IN t_number_table,
    p_position    IN BINARY_INTEGER
  )
  RETURN NUMBER
  IS
  BEGIN
    IF p_current_row.EXISTS(p_position) THEN
      RETURN p_current_row(p_position);
    ELSE
      RETURN NULL;
    END IF;
  EXCEPTION
    WHEN OTHERS THEN
       raise_application_error(-20001,'get_current_row:number: p_id='||p_position||' '||sqlerrm);
  END get_current_row;   
  ------------------------------------------------------------------------------
  -- replace multiple occurences of ":" to the one
  -- :::: -> :
  FUNCTION rr(p_str IN VARCHAR2)
  RETURN VARCHAR2
  IS 
  BEGIN
    RETURN ltrim(rtrim(regexp_replace(p_str,'[:]+',':'),':'),':');
  END rr;
  ------------------------------------------------------------------------------   
  -- remove HTML-tags (only simple cases are supported),
  -- special invisible characters like \t, \n ets
  -- and conver them to the xml
  FUNCTION get_xmlval(p_str IN VARCHAR2)
  RETURN VARCHAR2
  IS
    v_tmp largevarchar2;
  BEGIN
    -- p_str can be encoded html-string 
    -- wee need first convert to text
    v_tmp := regexp_replace(p_str,'<(BR)\s*/*>',chr(13)||chr(10),1,0,'i');
    v_tmp := regexp_replace(v_tmp,'<[^<>]+>',' ',1,0,'i');
    -- https://community.oracle.com/message/14074217#14074217
    v_tmp := regexp_replace(v_tmp, '[^[:print:]'||chr(13)||chr(10)||chr(9)||']', ' ');
    v_tmp := utl_i18n.unescape_reference(v_tmp); 
    -- and finally encode them
    --v_tmp := substr(v_tmp,1,32000);        
    v_tmp := dbms_xmlgen.convert( substr( v_tmp, 1, 32000 ) );

    RETURN v_tmp;
  END get_xmlval;  
  ------------------------------------------------------------------------------
  FUNCTION intersect_arrays(p_one IN apex_application_global.vc_arr2,
                            p_two IN apex_application_global.vc_arr2)
  RETURN apex_application_global.vc_arr2
  IS    
    v_ret apex_application_global.vc_arr2;
  BEGIN    
    FOR i IN 1..p_one.count LOOP
       FOR b IN 1..p_two.count LOOP
         IF p_one(i) = p_two(b) THEN
            v_ret(v_ret.count + 1) := p_one(i);
           EXIT;
         END IF;
       END LOOP;        
    END LOOP;

    RETURN v_ret;
  END intersect_arrays;
  ------------------------------------------------------------------------------
  -- get list of column-aliases from 
  -- report's SQL-Query
  FUNCTION get_query_column_list
  RETURN apex_application_global.vc_arr2
  IS
   v_cur         INTEGER; 
   v_colls_count BINARY_INTEGER; 
   v_columns     apex_application_global.vc_arr2;
   v_desc_tab    dbms_sql.desc_tab2;
   v_sql         largevarchar2;   
  BEGIN
    v_cur := dbms_sql.open_cursor(2);     
    v_sql := apex_plugin_util.replace_substitutions(p_value => l_report.report.sql_query,p_escape => false);
    log(v_sql);
    dbms_sql.parse(v_cur,v_sql,dbms_sql.native);     
    dbms_sql.describe_columns2(v_cur,v_colls_count,v_desc_tab);    
    FOR i IN 1..v_colls_count LOOP
         IF upper(v_desc_tab(i).col_name) != 'APXWS_ROW_PK' THEN --skip internal primary key if need
           v_columns(v_columns.count + 1) := v_desc_tab(i).col_name;
           log('Query column = '||v_desc_tab(i).col_name);
         END IF;
    END LOOP;                 
   dbms_sql.close_cursor(v_cur);   
   RETURN v_columns;
  EXCEPTION
    WHEN OTHERS THEN
      IF dbms_sql.is_open(v_cur) THEN
        dbms_sql.close_cursor(v_cur);
      END IF;  
      raise_application_error(-20001,'get_query_column_list: '||sqlerrm);
  END get_query_column_list;  
  ------------------------------------------------------------------------------
  -- filter p_delimetered_column_list column list: leave only displayed nonbreak columns here
  -- convert the result to pl/sql table
  FUNCTION get_cols_as_table(
    p_delimetered_column_list     IN VARCHAR2,
    p_displayed_nonbreak_columns  IN apex_application_global.vc_arr2
  )
  RETURN apex_application_global.vc_arr2
  IS
  BEGIN
    RETURN intersect_arrays(apex_util.string_to_table(rr(p_delimetered_column_list)),p_displayed_nonbreak_columns);
  END get_cols_as_table;
  ------------------------------------------------------------------------------
  -- fill all report metadata
  PROCEDURE init_t_report(p_app_id       IN NUMBER,
                          p_page_id      IN NUMBER,
                          p_region_id    IN NUMBER)
  IS
    l_report_id     NUMBER;
    v_query_targets apex_application_global.vc_arr2;
    l_new_report    ir_report; 
  BEGIN
    l_report := l_new_report;    
    --get base report id    
    log('l_region_id='||p_region_id);

    l_report_id := apex_ir.get_last_viewed_report_id (
      p_page_id   => p_page_id,
      p_region_id => p_region_id
    );

    log('l_base_report_id='||l_report_id);    

    -- get ddata from table
    SELECT r.* 
    INTO l_report.ir_data       
    FROM apex_application_page_ir_rpt r
    WHERE application_id = p_app_id 
      AND page_id = p_page_id
      AND session_id = nv('APP_SESSION')
      AND application_user = v('APP_USER')
      AND base_report_id = l_report_id
      AND ROWNUM <2;

    log('l_report_id='||l_report_id);
    l_report_id := l_report.ir_data.report_id;                                                                 


    l_report.report := apex_ir.get_report(
      p_page_id        => p_page_id,
      p_region_id      => p_region_id
    );
   -- Some users have cases where apex_ir.get_report returned SQL queries
   -- does not 100% reflect the current status of the report.
   -- For such special cases, a list of columns will be filtered:
   -- Only columns that are really represented in SQL query are permitted.
    l_report.ir_data.report_columns := apex_util.table_to_string(get_cols_as_table(l_report.ir_data.report_columns,get_query_column_list));

    -- here 2 plSQL tables wll be filled:
    -- l_report.displayed_columns - columns to display as Excel-columns
    -- l_report.break_really_on   - colums for control breaks, 
    -- values from these columns will be dislayed in Excel as contol-break row 
    -- l_report.break_really_on can differ from l_report.ir_data.break_enabled_on
    -- because some control-break columns can be hidden 
    <<displayed_columns>>   -- include calculation columns                                   
    FOR i IN (SELECT column_alias,
                     report_label,
                     heading_alignment,
                     column_alignment,
                     column_type,
                     format_mask AS  computation_format_mask,
                     nvl(instr(':'||l_report.ir_data.report_columns||':',':'||column_alias||':'),0) column_order ,
                     nvl(instr(':'||l_report.ir_data.break_enabled_on||':',':'||column_alias||':'),0) break_column_order,
                     column_link
                FROM apex_application_page_ir_col
               WHERE application_id = p_app_id
                 AND page_id = p_page_id
                 AND region_id = p_region_id
                 AND display_text_as != 'HIDDEN' --l_report.ir_data.report_columns can include HIDDEN columns
                 AND instr(':'||l_report.ir_data.report_columns||':',':'||column_alias||':') > 0
              UNION
              SELECT computation_column_alias,
                     computation_report_label,
                     'center' AS heading_alignment,
                     'right' AS column_alignment,
                     computation_column_type,
                     computation_format_mask,
                     nvl(instr(':'||l_report.ir_data.report_columns||':',':'||computation_column_alias||':'),0) column_order,
                     nvl(instr(':'||l_report.ir_data.break_enabled_on||':',':'||computation_column_alias||':'),0) break_column_order,
                     NULL AS column_link
              FROM apex_application_page_ir_comp
              WHERE application_id = p_app_id
                AND page_id = p_page_id
                AND report_id = l_report_id
                AND instr(':'||l_report.ir_data.report_columns||':',':'||computation_column_alias||':') > 0
              ORDER BY  break_column_order ASC,column_order ASC)
    LOOP                 
      l_report.column_header_labels(i.column_alias) := i.report_label; 
      l_report.col_format_mask(i.column_alias) := i.computation_format_mask;
      l_report.header_alignment(i.column_alias) := i.heading_alignment; 
      l_report.column_alignment(i.column_alias) := i.column_alignment; 
      l_report.column_link(i.column_alias) := i.column_link;
     IF i.column_order > 0 THEN
        IF i.break_column_order = 0 THEN 
          --displayed column
          l_report.displayed_columns(l_report.displayed_columns.count + 1) := i.column_alias;
        ELSE  
          --break column
          l_report.break_really_on(l_report.break_really_on.count + 1) := i.column_alias;
        END IF;
      END IF;  

      log('column='||i.column_alias||' l_report.column_header_labels='||i.report_label);
      log('column='||i.column_alias||' l_report.col_format_mask='||i.computation_format_mask);
      log('column='||i.column_alias||' l_report.header_alignment='||i.heading_alignment);
      log('column='||i.column_alias||' l_report.column_alignment='||i.column_alignment);    
      log('column='||i.column_alias||' l_report.column_link='||i.column_link);    
    END LOOP displayed_columns;    

    -- calculate columns count with aggregation separately
    l_report.sum_columns_on_break := get_cols_as_table(l_report.ir_data.sum_columns_on_break,l_report.displayed_columns);  
    l_report.avg_columns_on_break := get_cols_as_table(l_report.ir_data.avg_columns_on_break,l_report.displayed_columns);  
    l_report.max_columns_on_break := get_cols_as_table(l_report.ir_data.max_columns_on_break,l_report.displayed_columns);  
    l_report.min_columns_on_break := get_cols_as_table(l_report.ir_data.min_columns_on_break,l_report.displayed_columns);  
    l_report.median_columns_on_break := get_cols_as_table(l_report.ir_data.median_columns_on_break,l_report.displayed_columns); 
    l_report.count_columns_on_break := get_cols_as_table(l_report.ir_data.count_columns_on_break,l_report.displayed_columns);  
    l_report.count_distnt_col_on_break := get_cols_as_table(l_report.ir_data.count_distnt_col_on_break,l_report.displayed_columns); 

    -- calculate total count of columns with aggregation
    l_report.agg_cols_cnt := l_report.sum_columns_on_break.count + 
                             l_report.avg_columns_on_break.count +
                             l_report.max_columns_on_break.count + 
                             l_report.min_columns_on_break.count +
                             l_report.median_columns_on_break.count +
                             l_report.count_columns_on_break.count +
                             l_report.count_distnt_col_on_break.count;

    log('l_report.report_columns='||rr(l_report.ir_data.report_columns));    
    log('l_report.break_on='||rr(l_report.ir_data.break_enabled_on));
    log('l_report.sum_columns_on_break='||rr(l_report.ir_data.sum_columns_on_break));
    log('l_report.avg_columns_on_break='||rr(l_report.ir_data.avg_columns_on_break));
    log('l_report.max_columns_on_break='||rr(l_report.ir_data.max_columns_on_break));
    log('l_report.min_columns_on_break='||rr(l_report.ir_data.min_columns_on_break));
    log('l_report.median_columns_on_break='||rr(l_report.ir_data.median_columns_on_break));
    log('l_report.count_columns_on_break='||rr(l_report.ir_data.count_columns_on_break));    
    log('l_report.count_distnt_col_on_break='||rr(l_report.ir_data.count_distnt_col_on_break));
    log('l_report.break_really_on='||apex_util.table_to_string(l_report.break_really_on));
    log('l_report.agg_cols_cnt='||l_report.agg_cols_cnt);

    -- fill data for highlights
    -- v_query_targets is a list of columns needed to be added to calculate highlight

    FOR c IN cur_highlight(
      p_report_id => l_report_id,
      p_delimetered_column_list => l_report.ir_data.report_columns
    ) 
    LOOP
        IF c.highlight_row_color IS NOT NULL OR c.highlight_row_font_color IS NOT NULL THEN
          --is row highlight
          l_report.row_highlight(l_report.row_highlight.count + 1) := c;        
        ELSE
          l_report.col_highlight(l_report.col_highlight.count + 1) := c;           
        END IF;  
        v_query_targets(v_query_targets.count + 1) := c.condition_sql||' as HLIGHTS_'||(v_query_targets.count + 1);
    END LOOP;    

    -- original APEX Intractive Report calculates highlights by displaying,
    -- not by querying the data
    -- for me it was to difficult and I have decided to modify
    -- original SQL-query by adding highlights-columns

    IF v_query_targets.count  > 0 THEN
      -- uwr485kv is a random name 
      l_report.report.sql_query := 'SELECT '||apex_util.table_to_string(v_query_targets,','||chr(10))||', uwr485kv.* from ('||l_report.report.sql_query||') uwr485kv';
    END IF;
    -- plugin is always using the  restriction by maximal rows count
    IF instr(l_report.report.sql_query,':APXWS_MAX_ROW_CNT') = 0 THEN
      l_report.report.sql_query := 'SELECT * FROM ('|| l_report.report.sql_query ||') where rownum <= :APXWS_MAX_ROW_CNT';
    END IF;  

    -- the final LIST OF SQL-QUERY-COLUMNS:
    -- [columns for row/column-highlighting],      -- l_report.row_highlight.count + l_report.col_highlight.count
    -- [apxws_row_pk],                             -- l_report.skipped_columns, optional
    -- [columns for control break],                -- l_report.break_really_on.count
    -- [normal columns displayed as Excel-columns, 
    --  including computation-columns]             -- l_report.displayed_columns.count
    -- [aggregation results: sum]                  -- l_report.sum_columns_on_break.count
    -- [aggregation results: avg]                  -- l_report.avg_columns_on_break.count
    -- [aggregation results: max]                  -- l_report.max_columns_on_break.count
    -- [aggregation results: min]                  -- l_report.min_columns_on_break.count
    -- [aggregation results: median]               -- l_report.median_columns_on_break.count
    -- [aggregation results: count]                -- l_report.count_columns_on_break.count
    -- [aggregation results: count distinct]       -- l_report.count_distnt_col_on_break.count

    log('l_report.report.sql_query='||chr(10)||l_report.report.sql_query||chr(10));
  EXCEPTION
    WHEN no_data_found THEN
      raise_application_error(-20001,'No Interactive Report found on Page='||p_page_id||' Application='||p_app_id||' Please make sure that the report was running at least once by this session.');
  END init_t_report;  
  ------------------------------------------------------------------------------
  -- compare two rows to identify a case of contol-break
  FUNCTION is_control_break(
    p_curr_row  IN apex_application_global.vc_arr2,
    p_prev_row  IN apex_application_global.vc_arr2
  )
  RETURN BOOLEAN
  IS
    v_start_with      BINARY_INTEGER;
    v_end_with        BINARY_INTEGER;    
  BEGIN
    IF nvl(l_report.break_really_on.count,0) = 0  THEN
      RETURN false; --no control break
    END IF;
    -- see strucure of SQL-query in init_t_report-procedure
    v_start_with := 1 + 
                    l_report.skipped_columns + 
                    l_report.row_highlight.count + 
                    l_report.col_highlight.count;    
    v_end_with   := l_report.skipped_columns + 
                    l_report.row_highlight.count + 
                    l_report.col_highlight.count +
                    nvl(l_report.break_really_on.count,0);
    FOR i IN v_start_with..v_end_with LOOP
      IF p_curr_row(i) != p_prev_row(i) THEN
        RETURN true;
      END IF;
    END LOOP;
    RETURN false;
  END is_control_break;
  ------------------------------------------------------------------------------
  -- like to_char but with additional checking
  FUNCTION get_formatted_number(
    p_number         IN NUMBER,
    p_format_string  IN VARCHAR2,
    p_nls            IN VARCHAR2 DEFAULT NULL
  )
  RETURN VARCHAR2
  IS
    v_str VARCHAR2(100);
  BEGIN
    v_str := trim(TO_CHAR(p_number,p_format_string,p_nls));
    IF instr(v_str,'#') > 0 AND ltrim(v_str,'#') IS NULL THEN --format fail
      RAISE invalid_number;
    ELSE
      RETURN v_str;
    END IF;    
  END get_formatted_number;  
  ------------------------------------------------------------------------------
  -- t_cell_data saves information about Excel-cell
  --
  -- get_cell[date] tries to convert p_date_value
  -- as Excel-Date-Cell
  -- if not possible p_text_value is used to 
  -- convert as Excel-Text-Cell 
  FUNCTION get_cell(
    p_text_value  IN VARCHAR2,
    p_format_mask IN VARCHAR2,
    p_date_value  IN DATE
  )
  RETURN t_cell_data
  IS
    v_data t_cell_data;
  BEGIN     
     v_data.value := get_formatted_number(p_date_value - TO_DATE('01-03-1900','DD-MM-YYYY') + 61,'9999999999999990D00000000','NLS_NUMERIC_CHARACTERS = ''.,''');     
     v_data.datatype := 'DATE';

     -- https://github.com/glebovpavel/IR_to_MSExcel/issues/16
     -- thanks Valentine Nikitsky 
     IF p_format_mask IS NOT NULL THEN
       IF upper(p_format_mask)='SINCE' THEN
            v_data.text := apex_util.get_since(p_date_value);
            v_data.value := NULL;
            v_data.datatype := 'STRING';
       ELSE
          v_data.text := to_char(p_date_value,p_format_mask);  -- normally v_data.text used in XML only
       END IF;
     ELSE
       v_data.text := p_text_value;
     END IF;

     RETURN v_data;
  EXCEPTION
    WHEN invalid_number OR format_error THEN 
      v_data.value := NULL;          
      v_data.datatype := 'STRING';
      v_data.text := p_text_value;
      RETURN v_data;
  END get_cell;  
  ------------------------------------------------------------------------------
  -- t_cell_data saves information about Excel-cell
  --
  -- get_cell[number] tries to convert p_number_value
  -- as Excel-Number-Cell
  -- if not possible p_text_value is used to 
  -- convert as Excel-Text-Cell 
  FUNCTION get_cell(
    p_text_value   IN VARCHAR2,
    p_format_mask  IN VARCHAR2,
    p_number_value IN NUMBER
  )
  RETURN t_cell_data
  IS
    v_data t_cell_data;
  BEGIN
   v_data.datatype := 'NUMBER';   
   v_data.value := get_formatted_number(p_number_value,'9999999999999990D00000000','NLS_NUMERIC_CHARACTERS = ''.,''');

   IF p_format_mask IS NOT NULL THEN
     v_data.text := get_formatted_number(p_text_value,p_format_mask);
   ELSE
     v_data.text := p_text_value;
   END IF;

   RETURN v_data;
  EXCEPTION
    WHEN invalid_number OR conversion_error THEN 
      v_data.value := NULL;          
      v_data.datatype := 'STRING';
      v_data.text := p_text_value;
      RETURN v_data;
  END get_cell;  
  ------------------------------------------------------------------------------
  -- create a string with a list of all Contol-Break Headers.
  -- This string will be displayed as a row in Excel
  FUNCTION print_control_break_header_obj(p_current_row IN apex_application_global.vc_arr2) 
  RETURN VARCHAR2
  IS
    v_cb_xml  largevarchar2;
  BEGIN
    IF nvl(l_report.break_really_on.count,0) = 0  THEN
      RETURN ''; --no control break
    END IF;

    -- see strucure of SQL-query in init_t_report-procedure
    <<break_columns>>
    FOR i IN 1..nvl(l_report.break_really_on.count,0) LOOP
      --TODO: Add column header
      v_cb_xml := v_cb_xml ||
                  get_column_header_label(l_report.break_really_on(i))||': '||
                  get_current_row(p_current_row,i + 
                                                l_report.skipped_columns + 
                                                l_report.row_highlight.count + 
                                                l_report.col_highlight.count
                                 )||',';
    END LOOP visible_columns;

    RETURN  rtrim(v_cb_xml,',');
  END print_control_break_header_obj;
  ------------------------------------------------------------------------------
  -- find a position of given column in the list columns
  FUNCTION find_rel_position (
    p_curr_col_name    IN VARCHAR2,
    p_agg_rows         IN apex_application_global.vc_arr2
  )
  RETURN BINARY_INTEGER
  IS
    v_relative_position BINARY_INTEGER;
  BEGIN
    <<aggregated_rows>>
    FOR i IN 1..p_agg_rows.count LOOP
      IF p_curr_col_name = p_agg_rows(i) THEN        
         RETURN i;
      END IF;
    END LOOP aggregated_rows;

    RETURN NULL;
  END find_rel_position;
  ------------------------------------------------------------------------------
  FUNCTION get_agg_text(
    p_curr_col_name          IN VARCHAR2,
    p_agg_rows               IN apex_application_global.vc_arr2,
    p_current_row            IN apex_application_global.vc_arr2,
    p_agg_text               IN VARCHAR2,
    p_start_position_sql     IN BINARY_INTEGER, --start position in sql-query
    p_col_number             IN BINARY_INTEGER, --column position when displayed
    p_default_format_mask    IN VARCHAR2 DEFAULT NULL,
    p_overwrite_format_mask  IN VARCHAR2 DEFAULT NULL
  ) --should be used forcibly 
  RETURN VARCHAR2
  IS
    v_relative_position_sql  BINARY_INTEGER;   
    v_format_mask   apex_application_page_ir_comp.computation_format_mask%TYPE;
    v_agg_value     largevarchar2;
    v_row_value     largevarchar2;
    v_g_format_mask VARCHAR2(100);  
    v_col_alias     VARCHAR2(255);
  BEGIN
      -- see strucure of SQL-query in init_t_report-procedure
      --
      -- The aggregate values can see like 
      -- SELECT COL1,
      --        COL2,
      --        COL3,
      --        COL4,
      --        COL5,
      --       ,SUM(COL1) OVER()
      --       ,SUM(COL2) OVER()
      --       ,SUM(COL3) OVER()
      --       ,MIN(COL1) OVER()
      --       ,MAX(COL5) OVER()

      -- aggregate(s) of each type can be defined for many columns.
      -- v_relative_position shows the column position inside aggregate of each type
      -- for example above, 
      --              for SUM and COL1 v_relative_position will be 1, p_start_position_sql should be 5
      --              for SUM and COL2 v_relative_position will be 2  p_start_position_sql should be 5                
      --              for SUM and COL3 v_relative_position will be 3  p_start_position_sql should be 5
      -- after get_agg_text will be called for MIN
      --              for MIN and COL1 v_relative_position will be 1 p_start_position_sql should be 8
      -- p_start_position_sql should be   
      v_relative_position_sql := find_rel_position (p_curr_col_name,p_agg_rows); 
      IF v_relative_position_sql IS NOT NULL THEN
        v_col_alias := get_column_alias_sql(p_col_number);
        v_g_format_mask :=  get_col_format_mask(v_col_alias);   
        v_format_mask := nvl(v_g_format_mask,p_default_format_mask);
        v_format_mask := nvl(p_overwrite_format_mask,v_format_mask);
        v_row_value :=  get_current_row(p_current_row,p_start_position_sql + v_relative_position_sql);
        v_agg_value := trim(TO_CHAR(v_row_value,v_format_mask));

        RETURN  get_xmlval(p_agg_text||v_agg_value||' '||chr(10));
      ELSE
        RETURN  '';
      END IF;    
  EXCEPTION
     WHEN OTHERS THEN
        log('!Exception in get_agg_text');
        log('p_col_number='||p_col_number);
        log('v_col_alias='||v_col_alias);
        log('v_g_format_mask='||v_g_format_mask);
        log('p_default_format_mask='||p_default_format_mask);
        log('v_relative_position_sql='||v_relative_position_sql);
        log('p_start_position_sql='||p_start_position_sql);
        log('v_row_value='||v_row_value);
        log('v_format_mask='||v_format_mask);
        RAISE;
  END get_agg_text;
  ------------------------------------------------------------------------------
  -- Show all aggregates in Excel
  FUNCTION get_aggregate(p_current_row IN apex_application_global.vc_arr2) 
  RETURN apex_application_global.vc_arr2
  IS
    v_aggregate_xml   largevarchar2;
    v_agg_obj  apex_application_global.vc_arr2;
    v_position        BINARY_INTEGER;    
    v_i NUMBER := 0;
  BEGIN
    IF l_report.agg_cols_cnt  = 0 THEN      
      RETURN v_agg_obj;
    END IF;    

    <<visible_columns>>
    FOR i IN l_report.start_with..l_report.end_with LOOP
      v_aggregate_xml := '';
      v_i := v_i + 1;
      v_position := l_report.end_with; --aggregate are placed after displayed columns and computations      
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.sum_columns_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => ' ',
                                       p_start_position_sql      => v_position,
                                       p_col_number    => i);
      v_position := v_position + l_report.sum_columns_on_break.count;
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.avg_columns_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => 'Avgerage:',
                                       p_start_position_sql      => v_position,
                                       p_col_number    => i,
                                       p_default_format_mask   => '999G999G999G999G990D000');
      v_position := v_position + l_report.avg_columns_on_break.count;                                       
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.max_columns_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => 'Max:',
                                       p_start_position_sql      => v_position,
                                       p_col_number    => i);
      v_position := v_position + l_report.max_columns_on_break.count;                                 
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.min_columns_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => 'Min:',
                                       p_start_position_sql      => v_position,
                                       p_col_number    => i);
      v_position := v_position + l_report.min_columns_on_break.count;                                 
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.median_columns_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => 'Median:',
                                       p_start_position_sql      => v_position,
                                       p_col_number    => i,
                                       p_default_format_mask   => '999G999G999G999G990D000');
      v_position := v_position + l_report.median_columns_on_break.count;                                 
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.count_columns_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => 'Count:',
                                       p_start_position_sql      => v_position,
                                       p_col_number    => i,
                                       p_overwrite_format_mask   => '999G999G999G999G990');
      v_position := v_position + l_report.count_columns_on_break.count;                                 
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.count_distnt_col_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => 'Count distinct:',
                                       p_start_position_sql      => v_position,
                                       p_col_number    => i,
                                       p_overwrite_format_mask   => '999G999G999G999G990');
      v_agg_obj(v_i) := v_aggregate_xml;
    END LOOP visible_columns;
    RETURN  v_agg_obj;
  END get_aggregate;        
  ------------------------------------------------------------------------------
  -- chech that format string is valid
  FUNCTION can_show_as_date(p_format_string IN VARCHAR2)
  RETURN BOOLEAN
  IS 
    v_dummy VARCHAR2(50);
  BEGIN
    v_dummy := TO_CHAR(SYSDATE,p_format_string);
    RETURN true;
  EXCEPTION
    WHEN invalid_number OR format_error OR date_format_error OR conversion_error THEN 
      RETURN false;
  END can_show_as_date;    

  ------------------------------------------------------------------------------
  -- Excel has only DATE-format mask (no timezones etc)
  -- if format mask can be shown in excel - show column as date- type
  -- else - as string 
  PROCEDURE prepare_col_format_mask(
    p_col_number IN BINARY_INTEGER,
    p_data_type  IN BINARY_INTEGER
  )
  IS
    v_format_mask          t_formatmask;
    v_default_format_mask  t_formatmask;
    v_final_format_mask    t_formatmask;
    v_col_alias            VARCHAR2(255);    
  BEGIN
     v_default_format_mask := get_current_format(p_data_type);
     log('v_default_format_mask='||v_default_format_mask);
     BEGIN
       v_col_alias   := get_column_alias_sql(p_col_number);
       v_format_mask := get_col_format_mask(v_col_alias);
       log('v_col_alias='||v_col_alias||' v_format_mask='||v_format_mask);
     EXCEPTION
       WHEN OTHERS THEN 
          v_format_mask := '';
     END;  

     v_final_format_mask := nvl(v_format_mask,v_default_format_mask);
     IF can_show_as_date(v_final_format_mask) THEN
        log('Can show as date');
        IF v_col_alias IS NOT NULL THEN 
           set_col_format_mask(p_column_alias => v_col_alias,
                               p_format_mask  => v_final_format_mask);
        END IF;                       
        l_report.column_data_types(p_col_number) := 'DATE';
     ELSE
        log('Can not show as date');
        l_report.column_data_types(p_col_number) := 'STRING';
     END IF;
  END prepare_col_format_mask;
  ------------------------------------------------------------------------------
  -- Plugin tries to set column width 
  -- it measures the column width in html in the browser
  -- and send it to the plugin -- in p_width_str - parameter.
  -- p_coefficient is used to convert width in px (HTML) to the
  -- with used by Excel
  PROCEDURE print_column_widths(
      p_clob          IN OUT NOCOPY CLOB,
      p_buffer        IN OUT NOCOPY VARCHAR2,
      p_width_str     IN VARCHAR2,
      p_coefficient   IN NUMBER,
      p_custom_width  IN VARCHAR2
  )
  IS
    v_column_alias         apex_application_page_ir_col.column_alias%TYPE;
    v_i                    NUMBER := 0;
    v_col_width            NUMBER;
    v_custom_col_width     VARCHAR2(300);
    a_col_name_plus_width  apex_application_global.vc_arr2;    
    v_coefficient          NUMBER;    
  BEGIN  
    log('p_custom_width='||p_custom_width);
    add(p_clob,p_buffer,'<cols>'||chr(10));    
    -- save custom column widhts as pl/sql table    
    a_col_name_plus_width := apex_util.string_to_table(rtrim(p_width_str,','),',');    

    v_coefficient := nvl(p_coefficient,width_coefficient);
    IF v_coefficient = 0 THEN
      v_coefficient := width_coefficient;
    END IF;     

    FOR i IN 1..l_report.displayed_columns.count   LOOP
       v_column_alias := get_column_alias(i);
      -- if current column is not control break column
      -- !!ToDo: clear, why  why break_on is used instead of break_really_on ???
      IF apex_plugin_util.get_position_in_list(l_report.break_on,v_column_alias) IS NULL THEN 
        v_i := v_i + 1;
        BEGIN             
          -- check that column has custom width defined in plugin settings
          v_custom_col_width := regexp_substr(','||p_custom_width||',',','||v_column_alias||':\d+,');          
          IF v_custom_col_width IS NOT NULL THEN            
            v_col_width := to_number(ltrim(regexp_substr(v_custom_col_width,':\d+'),':'));
            log('v_col_width='||v_col_width||' for column '||v_column_alias);
          ELSE        
            v_col_width  := round(to_number(a_col_name_plus_width(v_i))/ v_coefficient);          
          END IF;
        EXCEPTION
          WHEN OTHERS THEN             
            v_col_width := -1;  
        END;

        IF v_col_width >= 0 THEN
          add(p_clob,p_buffer,'<col min="'||v_i||'" max="'||v_i||'" width="'||v_col_width||'" customWidth="1" />'||chr(10));        
        ELSE
          add(p_clob,p_buffer,'<col min="'||v_i||'" max="'||v_i||'" width="10" customWidth="0" />'||chr(10));        
        END IF;
      END IF;  
    END LOOP column_widths;
    v_i := 0;
    add(p_clob,p_buffer,'</cols>'||chr(10));
  END print_column_widths;

  ------------------------------------------------------------------------------
  -- links in Excel are saved in separate structure 
  -- and are linked with Excel cells
  -- this procedure add link information
  -- to there structures
  PROCEDURE add_link(
    p_cell_addr IN VARCHAR2,
    p_link      IN VARCHAR2,
    p_row       IN apex_application_global.vc_arr2,
    p_links     IN OUT NOCOPY  apex_application_global.vc_arr2,
    p_links_ref IN OUT NOCOPY  apex_application_global.vc_arr2
  )
  IS
    v_substitution   t_large_varchar2;
    v_link           t_large_varchar2;
    v_cnt             NUMBER DEFAULT 100;

    FUNCTION get_subtitution_value(p_substitution IN VARCHAR2)
    RETURN VARCHAR2
    IS
    BEGIN
      RETURN p_row(l_report.all_columns(replace(p_substitution,'#','')));
    EXCEPTION
      WHEN OTHERS THEN
        RETURN '';
    END get_subtitution_value;

  BEGIN
    IF p_link IS NULL OR p_links.count >= 65530 THEN
      return; --EXCEL LINKS LIMIT
    END IF;  
    v_link := p_link;

    v_substitution := regexp_substr(v_link,'#[^#]+#');
    LOOP
     EXIT WHEN v_substitution IS NULL;
     v_link := replace(v_link,v_substitution,get_subtitution_value(v_substitution));
     v_substitution := regexp_substr(v_link,'#[^#]+#');       
    END LOOP;
    IF substr(v_link,1,4) = 'f?p=' THEN
        v_link := owa_util.get_cgi_env ('REQUEST_PROTOCOL') || '://'  
               || owa_util.get_cgi_env ('HTTP_HOST')        || ':'  
               || owa_util.get_cgi_env ('SERVER_PORT')      || '/'  
               || ltrim(owa_util.get_cgi_env ('SCRIPT_NAME'),'/')|| '/'
               || ltrim(apex_plugin_util.replace_substitutions(v_link),'/');
    ELSE
      v_link := apex_plugin_util.replace_substitutions(v_link);
    END IF;

    log('v_link='||v_link);
    p_links(p_links.count + 1) := '<hyperlink ref="'||p_cell_addr||'" r:id="rId'||(p_links.count+1)||'" />';
    p_links_ref(p_links_ref.count + 1) := '<Relationship Id="rId'||p_links.count||'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="'||dbms_xmlgen.convert(v_link)||'" TargetMode="External"/>';      
  END add_link;
  ------------------------------------------------------------------------------
  -- generate all XML-CLOBS needed tobuild an Excle-file
  -- from report-data
  PROCEDURE generate_from_report(
    v_clob          IN OUT NOCOPY CLOB,
    v_strings_clob  IN OUT NOCOPY CLOB,
    v_links_clob    IN OUT NOCOPY CLOB,
    p_width_str     IN VARCHAR2,
    p_coefficient   IN NUMBER,
    p_max_rows      IN INTEGER,
    p_autofilter    IN CHAR,
    p_export_links  IN CHAR DEFAULT 'N',
    p_custom_width  IN VARCHAR2
  ) 
  IS
    v_cur                INTEGER; 
    v_result             INTEGER;
    v_colls_count        BINARY_INTEGER;
    v_row                apex_application_global.vc_arr2;
    v_date_row           t_date_table;
    v_number_row         t_number_table;
    v_char_dummy         VARCHAR2(1);
    v_date_dummy         DATE;
    v_number_dummy       NUMBER;
    v_prev_row           apex_application_global.vc_arr2;
    v_columns            apex_application_global.vc_arr2;
    v_current_row        NUMBER DEFAULT 0;   
    v_inside             BOOLEAN DEFAULT false;
    v_sql                largevarchar2;
    v_bind_variables     dbms_sql.varchar2_table;
    v_buffer             largevarchar2;
    v_bind_var_name      VARCHAR2(255);
    v_binded             BOOLEAN; 
    v_format_mask        t_formatmask;  
    v_rowset_count       NUMBER := 0; 
    v_i                 NUMBER := 0;
    v_column_alias      apex_application_page_ir_col.column_alias%TYPE;
    v_row_color         VARCHAR2(10); 
    v_row_back_color    VARCHAR2(10);
    v_cell_color        VARCHAR2(10);
    v_cell_back_color   VARCHAR2(10);     
    v_column_data_type       columntype;
    v_cell_data         t_cell_data;    


    v_strings           apex_application_global.vc_arr2;
    v_links             apex_application_global.vc_arr2; -- which cell has a link
    v_links_ref         apex_application_global.vc_arr2; -- URL itself
    v_rownum            BINARY_INTEGER DEFAULT 1;

    v_style_id          BINARY_INTEGER;         
    v_string_buffer     t_large_varchar2; 
    v_break_header      t_large_varchar2;
    v_links_buffer      t_large_varchar2;
    v_last_agg_obj      apex_application_global.vc_arr2;

    v_link              apex_application_page_ir_col.column_link%TYPE;
    v_cell_addr         VARCHAR2(12);

    -- print Excel-cell containing text - data into XML-CLOB (using buffer)
    PROCEDURE print_char_cell(
      p_cell_addr IN VARCHAR2,
      p_string    IN VARCHAR2,
      p_clob      IN OUT NOCOPY CLOB,
      p_buffer    IN OUT NOCOPY VARCHAR2,
      p_style_id  IN NUMBER DEFAULT NULL
     )
    IS
     v_style     VARCHAR2(20);     
    BEGIN
      IF p_style_id IS NOT NULL THEN
       v_style := ' s="'||p_style_id||'" ';
      END IF;      
      add(p_clob,p_buffer,'<c r="'||p_cell_addr||'" t="s" '||v_style||'>'||chr(10)
                         ||'<v>' || TO_CHAR(v_strings.count)|| '</v>'||chr(10)                 
                         ||'</c>'||chr(10));
      v_strings(v_strings.count + 1) := p_string;      
    END print_char_cell;

    -- print Excel-cell containing number into XML-CLOB (using buffer)
    PROCEDURE print_number_cell(
      p_cell_addr IN VARCHAR2,
      p_value     IN VARCHAR2,
      p_clob      IN OUT NOCOPY CLOB,
      p_buffer    IN OUT NOCOPY VARCHAR2,
      p_style_id  IN NUMBER DEFAULT NULL
     )
    IS
      v_style VARCHAR2(20);
    BEGIN
      IF p_style_id IS NOT NULL THEN
       v_style := ' s="'||p_style_id||'" ';
      END IF;
      add(p_clob,p_buffer,'<c r="'||p_cell_addr||'" '||v_style||'>'||chr(10)
                         ||'<v>'||p_value|| '</v>'||chr(10)
                         ||'</c>'||chr(10));

    END print_number_cell; 

    -- print Excel-cell containing aggregates - as text -  into XML-CLOB (using buffer)
    PROCEDURE print_agg(
      p_agg_obj apex_application_global.vc_arr2,
      p_rownum    IN OUT BINARY_INTEGER,
      p_clob      IN OUT NOCOPY CLOB,
      p_buffer    IN OUT NOCOPY VARCHAR2
    )
    IS
        v_agg_clob         CLOB;
        v_agg_buffer       t_large_varchar2;
        v_agg_strings_cnt  BINARY_INTEGER DEFAULT 1; 
    BEGIN
        dbms_lob.createtemporary(v_agg_clob,true);
        /*PRINTAGG*/
        IF p_agg_obj.last IS NOT NULL THEN
            dbms_lob.trim(v_agg_clob,0);
            v_agg_buffer := '';       
            v_agg_strings_cnt := 1;       

            FOR y IN p_agg_obj.first..p_agg_obj.last
            LOOP
                v_agg_strings_cnt := greatest(length(regexp_replace('[^:]','')) + 1,v_agg_strings_cnt);
                v_style_id := get_aggregate_style_id(p_font         => '',
                                                    p_back         => '',
                                                    p_data_type    => 'CHAR',
                                                    p_format_mask  => '');
                print_char_cell(p_cell_addr  => get_cell_name(y,v_rownum),
                                p_string     => get_xmlval( rtrim( p_agg_obj(y),chr(10))),
                                p_clob       => v_agg_clob,
                                p_buffer     =>  v_agg_buffer,           
                                p_style_id   => v_style_id);        
            END LOOP;
            add(p_clob,v_buffer,'<row ht="'||(v_agg_strings_cnt * string_height)||'">'||chr(10),true);
            add(v_agg_clob,v_agg_buffer,' ',true);
            dbms_lob.copy( dest_lob => p_clob,
                            src_lob => v_agg_clob,
                            amount => dbms_lob.getlength(v_agg_clob),
                            dest_offset => dbms_lob.getlength(p_clob),
                            src_offset => 1);
            add(p_clob,v_buffer,'</row>'||chr(10));
            p_rownum := p_rownum + 1;
        END IF;
        dbms_lob.freetemporary(v_agg_clob);
  END print_agg;
  --


  --
  BEGIN  -- BEGIN of the MAIN-function generate_from_report   
     -- inline often used functions
     PRAGMA inline(add,'YES');
     PRAGMA inline(log,'YES');  
     -- In earlier versions code to get data from the report
     -- was splitted from the code to generate Excel.
     -- But it was too slow, after combining it
     -- the readability was worsened
     -- but performance was increased up to x10

    add(v_clob,v_buffer,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'||chr(10));
    add(v_clob,v_buffer,'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <dimension ref="A1"/>
        <sheetViews>
            <sheetView tabSelected="1" workbookViewId="0">
            <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
            <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
            </sheetView>
            </sheetViews>
        <sheetFormatPr baseColWidth="10" defaultColWidth="10" defaultRowHeight="15"/>'||chr(10),true);

    print_column_widths(
      p_clob          => v_clob,
      p_buffer        => v_buffer,
      p_width_str     => p_width_str,
      p_coefficient   => p_coefficient,
      p_custom_width  => p_custom_width
    );    

    add(v_clob,v_buffer,'<sheetData>'||chr(10));
     --column header
    add(v_clob,v_buffer,'<row>'||chr(10));      

    v_cur := dbms_sql.open_cursor(2); 
    v_sql := apex_plugin_util.replace_substitutions(
      p_value  => l_report.report.sql_query,
      p_escape => false
    );    
    dbms_sql.parse(v_cur,v_sql,dbms_sql.native);     
    dbms_sql.describe_columns2(v_cur,v_colls_count,l_report.desc_tab);    
    --skip internal primary key if need
    FOR i IN 1..l_report.desc_tab.count LOOP            
      IF lower(l_report.desc_tab(i).col_name) = 'apxws_row_pk' THEN
        l_report.skipped_columns := 1;
      END IF;      
    END LOOP;
    -- see strucure of SQL-query in init_t_report-procedure
    l_report.start_with := 1 + 
                           l_report.skipped_columns +
                           nvl(l_report.break_really_on.count,0) + 
                           l_report.row_highlight.count + 
                           l_report.col_highlight.count;
    l_report.end_with   := l_report.skipped_columns + 
                           nvl(l_report.break_really_on.count,0) + 
                           l_report.row_highlight.count + 
                           l_report.col_highlight.count +
                           l_report.displayed_columns.count;    

    log('l_report.start_with='||l_report.start_with);
    log('l_report.end_with='||l_report.end_with);
    log('l_report.skipped_columns='||l_report.skipped_columns);
    -- init column datatypes and format masks
    FOR i IN 1..l_report.desc_tab.count LOOP      
      log('column_type='||l_report.desc_tab(i).col_type);
      -- don't know why column types in dbms_sql.describe_columns2 do not correspond to types in dbms_types      
      IF l_report.desc_tab(i).col_type IN (dbms_types.typecode_timestamp_tz,181,
                                           dbms_types.typecode_timestamp_ltz,231,
                                           dbms_types.typecode_timestamp,180,
                                           dbms_types.typecode_date) 
      THEN
         prepare_col_format_mask(
           p_col_number => i,
           p_data_type => l_report.desc_tab(i).col_type
         );      
      ELSIF l_report.desc_tab(i).col_type = dbms_types.typecode_number THEN
         l_report.column_data_types(i) := 'NUMBER';
      ELSE
         l_report.column_data_types(i) := 'STRING';
         log('column_type='||l_report.desc_tab(i).col_type||' STRING');
      END IF;     
      l_report.all_columns(upper(l_report.desc_tab(i).col_name)) := i;
    END LOOP;

    v_bind_variables := wwv_flow_utilities.get_binds(v_sql);

    <<headers>>
    FOR i IN 1..l_report.displayed_columns.count   LOOP
       v_column_alias := get_column_alias(i);
      -- if current column is not control break column
      -- !! to do check why break_on is used
      IF apex_plugin_util.get_position_in_list(l_report.break_on,v_column_alias) IS NULL THEN       
        print_char_cell(
          p_cell_addr  => get_cell_name(i,v_rownum),
          p_string     => get_xmlval(get_column_header_label(v_column_alias)),
          p_clob       => v_clob,
          p_buffer     => v_buffer,
          p_style_id   => get_header_style_id(p_align  => get_header_alignment(v_column_alias))
        );        
      END IF;  
    END LOOP headers;    
    v_rownum := v_rownum + 1;
    add(v_clob,v_buffer,'</row>'||chr(10)); 

    log('<<bind variables>>');
    <<bind_variables>>
    FOR i IN 1..v_bind_variables.count LOOP      
      v_bind_var_name := ltrim(v_bind_variables(i),':');
      IF v_bind_var_name = 'APXWS_MAX_ROW_CNT' THEN      
         -- remove max_rows
         dbms_sql.bind_variable (v_cur,'APXWS_MAX_ROW_CNT',p_max_rows);    
         log('Bind variable ('||i||')'||v_bind_var_name||'<'||p_max_rows||'>');
      ELSE
        v_binded := false; 
        --first look report bind variables (filtering, search etc)
        <<bind_report_variables>>
        FOR a IN 1..l_report.report.binds.count LOOP
          IF v_bind_var_name = l_report.report.binds(a).name THEN
             dbms_sql.bind_variable (v_cur,v_bind_var_name,l_report.report.binds(a).value);
             log('Bind variable as report variable ('||i||')'||v_bind_var_name||'<'||l_report.report.binds(a).value||'>');
             v_binded := true;
             EXIT;
          END IF;
        END LOOP bind_report_variables;
        -- substantive strings in sql-queries can have bind variables too
        -- these variables are not in v_report.binds
        -- and need to be binded separately
        IF NOT v_binded THEN
          dbms_sql.bind_variable (v_cur,v_bind_var_name,v(v_bind_var_name));
          log('Bind variable ('||i||')'||v_bind_var_name||'<'||v(v_bind_var_name)||'>');          
        END IF;        
       END IF; 
    END LOOP;          

    log('<<define_columns>>');    
    FOR i IN 1..v_colls_count LOOP
       log('define column '||i);
       log('column type '||l_report.desc_tab(i).col_type);      
       IF l_report.column_data_types(i) = 'DATE' THEN   
         dbms_sql.define_column(v_cur, i, v_date_dummy);
       ELSIF l_report.column_data_types(i) = 'NUMBER' THEN   
         dbms_sql.define_column(v_cur, i, v_number_dummy);
       ELSE --STRING
         dbms_sql.define_column(v_cur, i, v_char_dummy,32767);
       END IF;         
    END LOOP define_columns;    

    v_result := dbms_sql.execute(v_cur);         

    log('<<main_cycle>>');
    <<main_cycle>>
    LOOP 
         IF dbms_sql.fetch_rows(v_cur)>0 THEN          
         log('<<fetch>>');
           -- get column values of the row 
            v_current_row := v_current_row + 1;
            <<query_columns>>
            FOR i IN 1..v_colls_count LOOP               
               log('column type '||l_report.desc_tab(i).col_type);
               v_row(i) := ' ';
               v_date_row(i) := NULL;
               v_number_row(i) := NULL;
               IF l_report.column_data_types(i) = 'DATE' THEN
                 dbms_sql.column_value(v_cur, i,v_date_row(i));
                 v_row(i) := TO_CHAR(v_date_row(i));
               ELSIF l_report.column_data_types(i) = 'NUMBER' THEN
                dbms_sql.column_value(v_cur, i,v_number_row(i));
                v_row(i) := TO_CHAR(v_number_row(i));
               ELSE 
                 dbms_sql.column_value(v_cur, i,v_row(i));                 
               END IF;  
            END LOOP query_columns;     
            --check control break
            IF v_current_row > 1 THEN
             IF is_control_break(v_row,v_prev_row) THEN
                v_inside := false;                                                         
                print_agg(v_last_agg_obj,v_rownum,v_clob,v_buffer);                
             END IF;
            END IF;
            IF NOT v_inside THEN
                v_break_header :=  print_control_break_header_obj(v_row);
                IF v_break_header IS NOT NULL THEN
                    add(v_clob,v_buffer,'<row>'||chr(10));
                    print_char_cell(
                      p_cell_addr => get_cell_name(1,v_rownum),
                      p_string    => get_xmlval(v_break_header),
                      p_clob      => v_clob,
                      p_buffer    => v_buffer,
                      p_style_id  => get_header_style_id(p_back => NULL,p_align  => 'left')
                    );         
                    v_rownum := v_rownum + 1;
                    add(v_clob,v_buffer,'</row>'||chr(10));
                END IF;
               v_last_agg_obj := get_aggregate(v_row);                           
               v_inside := true;
            END IF;            --                        
            FOR i IN 1..v_colls_count LOOP
              v_prev_row(i) := v_row(i);                           
            END LOOP;

            add(v_clob,v_buffer,'<row>'||chr(10));
            /* CELLS INSIDE ROW PRINTING*/
            v_row_color := NULL;
            v_row_back_color  := NULL;
            <<row_highlights>>
            FOR h IN 1..l_report.row_highlight.count LOOP
              BEGIN 
                IF get_current_row(v_row,l_report.row_highlight(h).cond_number) IS NOT NULL THEN
                    v_row_color       := l_report.row_highlight(h).highlight_row_font_color;
                    v_row_back_color  := l_report.row_highlight(h).highlight_row_color;
                END IF;
              EXCEPTION       
                WHEN no_data_found THEN
                    log('row_highlights: ='||' end_with='||l_report.end_with||' agg_cols_cnt='||l_report.agg_cols_cnt||' COND_NUMBER='||l_report.row_highlight(h).cond_number||' h='||h);
              END; 
            END LOOP row_highlights;

            <<visible_columns>>
            v_i := 0;
            FOR i IN l_report.start_with..l_report.end_with LOOP
                v_i := v_i + 1;
                v_cell_color       := NULL;
                v_cell_back_color  := NULL;
                v_cell_data.value  := NULL;  
                v_cell_data.text   := NULL; 
                v_column_alias     := get_column_alias_sql(i);
                v_column_data_type := get_column_data_type(i);
                v_format_mask      := get_col_format_mask(v_column_alias);    
                IF  p_export_links = 'Y' THEN
                  v_link := get_col_link(v_column_alias);
                ELSE
                  v_link := '';
                END IF;
                v_cell_addr := get_cell_name(v_i,v_rownum);
                add_link(
                  p_cell_addr => v_cell_addr,
                  p_link => v_link,
                  p_row => v_row,
                  p_links => v_links,
                  p_links_ref => v_links_ref
                );

                IF v_column_data_type = 'DATE' THEN
                    v_cell_data := get_cell(get_current_row(v_row,i),v_format_mask,get_current_row(v_date_row,i));
                ELSIF  v_column_data_type = 'NUMBER' THEN      
                    v_cell_data := get_cell(get_current_row(v_row,i),v_format_mask,get_current_row(v_number_row,i));
                ELSE --STRING
                    v_format_mask := NULL;
                    v_cell_data.value  := NULL;  
                    v_cell_data.datatype := 'STRING';
                    v_cell_data.text   := get_current_row(v_row,i);
                END IF; 

                --check that cell need to be highlighted
                <<cell_highlights>>
                FOR h IN 1..l_report.col_highlight.count LOOP
                  BEGIN
                    IF get_current_row(v_row,l_report.col_highlight(h).cond_number) IS NOT NULL 
                        AND v_column_alias = l_report.col_highlight(h).condition_column_name 
                    THEN
                        v_cell_color       := l_report.col_highlight(h).highlight_cell_font_color;
                        v_cell_back_color  := l_report.col_highlight(h).highlight_cell_color;
                    END IF;
                  EXCEPTION
                    WHEN no_data_found THEN
                        log('col_highlights: ='||' end_with='||l_report.end_with||' agg_cols_cnt='||l_report.agg_cols_cnt||' COND_NUMBER='||l_report.col_highlight(h).cond_number||' h='||h); 
                  END;
                END LOOP cell_highlights;

                BEGIN                                
                  v_style_id := get_style_id(
                    p_font_color   => nvl(v_cell_color,v_row_color), -- cell highlight owerwrites row highlight 
                    p_back_color   => nvl(v_cell_back_color,v_row_back_color), -- cell highlight owerwrites row highlight 
                    p_data_type    => v_cell_data.datatype,
                    p_format_mask  => v_format_mask,
                    p_align        => lower(get_column_alignment(v_column_alias)),
                    p_is_link      => (v_link IS NULL)
                  );
                  IF v_cell_data.datatype IN ('NUMBER') THEN
                      print_number_cell(p_cell_addr => v_cell_addr,
                                        p_value     => v_cell_data.value,
                                        p_clob      => v_clob,
                                        p_buffer    => v_buffer,
                                        p_style_id  => v_style_id
                                      );

                  ELSIF  v_cell_data.datatype IN ('DATE') THEN
                      add(v_clob,v_buffer,'<c r="'||v_cell_addr||'"  s="'||v_style_id||'">'||chr(10)
                                          ||'<v>'|| v_cell_data.value|| '</v>'||chr(10)
                                          ||'</c>'||chr(10)
                          );
                  ELSE --STRING
                      print_char_cell(p_cell_addr => v_cell_addr,
                                      p_string    => get_xmlval(v_cell_data.text),
                                      p_clob      => v_clob,
                                      p_buffer    => v_buffer,
                                      p_style_id  => v_style_id
                                      );
                  END IF;       
                EXCEPTION
                  WHEN no_data_found THEN
                    NULL;
                END;
            END LOOP visible_columns;            

            add(v_clob,v_buffer,'</row>'||chr(10));         
            v_rownum := v_rownum + 1;
         ELSE
           EXIT; 
         END IF;
    END LOOP main_cycle;    
    IF v_inside THEN       
       v_inside := false;
       print_agg(v_last_agg_obj,v_rownum,v_clob,v_buffer);
    END IF;  
    dbms_sql.close_cursor(v_cur);  

    add(v_clob,v_buffer,'</sheetData>'||chr(10));
    IF p_autofilter = 'Y' THEN
        add(v_clob,v_buffer,'<autoFilter ref="A1:' || get_cell_name(v_colls_count,v_rownum-1) || '"/>');
    END IF;
    -- print links
    IF v_links.count >  0 THEN 
        add(v_clob,v_buffer,'<hyperlinks>'||chr(10));        
        FOR p IN 1..v_links.count LOOP     
           add(v_clob,v_buffer,v_links(p)||chr(10));
        END LOOP;
        add(v_clob,v_buffer,'</hyperlinks>'||chr(10));
    END IF;

    add(v_clob,v_buffer,'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'||chr(10),true);

    add(v_strings_clob,v_string_buffer,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'||chr(10));
    add(v_strings_clob,v_string_buffer,'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' || v_strings.count() || '" uniqueCount="' || v_strings.count() || '">'||chr(10));

    FOR i IN 1 .. v_strings.count() LOOP
      add(v_strings_clob,v_string_buffer,'<si><t>'||v_strings(i)|| '</t></si>'||chr(10));
    END LOOP; 
    add(v_strings_clob,v_string_buffer,'</sst>'||chr(10),true);

    -- print links in sheet1.xml.rels
    add(v_links_clob,v_links_buffer,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'||chr(10));
    add(v_links_clob,v_links_buffer,'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'||chr(10));
    FOR l IN 1..v_links_ref.count LOOP
      add(v_links_clob,v_links_buffer,v_links_ref(l)||chr(10));
    END LOOP;
    add(v_links_clob,v_links_buffer,'</Relationships>'||chr(10),true);
  END generate_from_report;
  ------------------------------------------------------------------------------

  PROCEDURE get_report(
     p_app_id            IN NUMBER,
     p_page_id           IN NUMBER,  
     p_region_id         IN NUMBER,
     p_max_rows          IN NUMBER,            -- maximum rows for export   
     v_clob              IN OUT NOCOPY CLOB,
     p_strings           IN OUT NOCOPY CLOB,
     p_links             IN OUT NOCOPY CLOB,
     p_col_length        IN VARCHAR2,
     p_width_coefficient IN NUMBER, 
     p_autofilter        IN CHAR DEFAULT 'Y',
     p_export_links      IN CHAR DEFAULT 'N',
     p_custom_width      IN VARCHAR2
  )                 
  IS  
  BEGIN
    log('p_app_id='||p_app_id);
    log('p_page_id='||p_page_id);
    log('p_region_id='||p_region_id);
    log('p_max_rows='||p_max_rows);
    log('p_col_length='||p_col_length);
    log('p_autofilter='||p_autofilter);    
    log('p_export_links='||p_export_links);

    init_t_report(p_app_id,p_page_id,p_region_id);
    generate_from_report(v_clob,p_strings,p_links,p_col_length,p_width_coefficient,p_max_rows,p_autofilter,p_export_links,p_custom_width);
  END get_report;   

  ------------------------------------------------------------------------------
  /* 
    function to handle cases of 'in' and 'not in' conditions for highlights
       used in cursor cur_highlight

    Author: Srihari Ravva
  */ 
  FUNCTION get_highlight_in_cond_sql(
    p_condition_expression  IN apex_application_page_ir_cond.condition_expression%TYPE,
    p_condition_sql         IN apex_application_page_ir_cond.condition_sql%TYPE,
    p_condition_column_name IN apex_application_page_ir_cond.condition_column_name%TYPE
  )
  RETURN VARCHAR2 
  IS
    v_condition_sql_tmp  VARCHAR2(32767);
    v_condition_sql      VARCHAR2(32767);
    v_arr_cond_expr      apex_application_global.vc_arr2;
    v_arr_cond_sql       apex_application_global.vc_arr2;    
  BEGIN
    v_condition_sql := replace(replace(p_condition_sql,'#APXWS_HL_ID#','1'),'#APXWS_CC_EXPR#','"'||p_condition_column_name||'"');
    v_condition_sql_tmp := substr(v_condition_sql,instr(v_condition_sql,'#'),instr(v_condition_sql,'#',-1)-instr(v_condition_sql,'#')+1);

    v_arr_cond_expr := apex_util.string_to_table(p_condition_expression,',');
    v_arr_cond_sql := apex_util.string_to_table(v_condition_sql_tmp,',');

    FOR i IN 1..v_arr_cond_expr.count
    LOOP
        -- consider everything as varchar2
        -- 'in' and 'not in' highlight conditions are not possible for DATE columns from IR
        v_condition_sql := replace(v_condition_sql,v_arr_cond_sql(i),''''||TO_CHAR(v_arr_cond_expr(i))||'''');
    END LOOP;
    RETURN v_condition_sql;
  END get_highlight_in_cond_sql;  

  ------------------------------------------------------------------------------  
  PROCEDURE add1file(
    p_zipped_blob IN OUT NOCOPY BLOB,
    p_name IN VARCHAR2,
    p_content IN CLOB
  )
  IS
    v_desc_offset PLS_INTEGER := 1;
    v_src_offset  PLS_INTEGER := 1;
    v_lang        PLS_INTEGER := 0;
    v_warning     PLS_INTEGER := 0;
    v_blob        BLOB;
  BEGIN
    dbms_lob.createtemporary(v_blob,true);
    dbms_lob.converttoblob(v_blob,p_content, dbms_lob.getlength(p_content), v_desc_offset, v_src_offset, dbms_lob.default_csid, v_lang, v_warning);
    as_zip.add1file( p_zipped_blob, p_name, v_blob);
    dbms_lob.freetemporary(v_blob);
  END add1file;  

  ------------------------------------------------------------------------------
  FUNCTION get_max_rows (p_app_id      IN NUMBER,
                         p_page_id     IN NUMBER,
                         p_region_id   IN NUMBER)
  RETURN NUMBER
  IS 
    v_max_row_count NUMBER;
  BEGIN
    SELECT max_row_count 
    INTO v_max_row_count
    FROM apex_application_page_ir
    WHERE application_id = p_app_id
      AND page_id = p_page_id
      AND region_id = p_region_id
      AND ROWNUM <2;

     RETURN v_max_row_count;
  END get_max_rows;   
  ------------------------------------------------------------------------------  
  FUNCTION get_file_name (p_app_id      IN NUMBER,
                          p_page_id     IN NUMBER,
                          p_region_id   IN NUMBER)
  RETURN VARCHAR2
  IS 
    v_filename VARCHAR2(255);
  BEGIN
    SELECT filename 
    INTO v_filename
    FROM apex_application_page_ir
    WHERE application_id = p_app_id
      AND page_id = p_page_id
      AND region_id = p_region_id
      AND ROWNUM <2;

     RETURN apex_plugin_util.replace_substitutions(nvl(v_filename,'Excel'));
  END get_file_name;   
  ------------------------------------------------------------------------------
  -- download binary file
  PROCEDURE download(p_data        IN OUT NOCOPY BLOB,
                     p_mime_type   IN VARCHAR2,
                     p_file_name   IN VARCHAR2)
  IS
  BEGIN
    owa_util.mime_header( 
      ccontent_type=> p_mime_type, 
      bclose_header => false
    );
    sys.htp.p('Content-Length: ' || dbms_lob.getlength( p_data ) );
    sys.htp.p('Content-disposition: attachment; filename='||p_file_name );    
    sys.htp.p('Cache-Control: must-revalidate, max-age=0');
    sys.htp.p('Expires: Thu, 01 Jan 1970 01:00:00 CET');
    sys.htp.p('Set-Cookie: GPV_DOWNLOAD_STARTED=1;');
    owa_util.http_header_close;    
    wpg_docload.download_file( p_data );  
  EXCEPTION
     WHEN OTHERS THEN 
        raise_application_error(-20001,'Download (blob)'||sqlerrm);
  END download;
  ------------------------------------------------------------------------------

  PROCEDURE download_excel(
    p_app_id       IN NUMBER,
    p_page_id      IN NUMBER,
    p_region_id    IN NUMBER, 
    p_col_length   IN VARCHAR2 DEFAULT NULL,
    p_max_rows     IN NUMBER,
    p_autofilter   IN CHAR DEFAULT 'Y',
    p_export_links IN CHAR DEFAULT 'N',
    p_custom_width IN VARCHAR2
  ) IS
    t_template BLOB;
    t_excel    BLOB;
    v_cells    CLOB;
    v_strings  CLOB;
    v_links    CLOB;
    zip_files  as_zip.file_list;
  BEGIN        
    dbms_lob.createtemporary(t_excel,true);    
    dbms_lob.createtemporary(v_cells,true);
    dbms_lob.createtemporary(v_strings,true);
    dbms_lob.createtemporary(v_links,true);    

    get_report(p_app_id,
               p_page_id ,  
               p_region_id,
               p_max_rows,            -- maximum rows for export   
               v_cells ,
               v_strings,
               v_links,
               p_col_length,
               width_coefficient,
               p_autofilter,
               p_export_links,
               p_custom_width
              );

    SELECT file_content
    INTO t_template
    FROM apex_appl_plugin_files 
    WHERE file_name = 'ExcelTemplate.zip'
      AND application_id = p_app_id
      AND plugin_name='AT.FRT.GPV_IR_TO_MSEXCEL';

    zip_files  := as_zip.get_file_list( t_template );
    FOR i IN zip_files.first() .. zip_files.last LOOP
      as_zip.add1file( t_excel, zip_files( i ), as_zip.get_file( t_template, zip_files( i ) ) );
    END LOOP;    

    add1file( t_excel, 'xl/styles.xml', get_styles_xml);    
    add1file( t_excel, 'xl/worksheets/Sheet1.xml', v_cells);
    add1file( t_excel, 'xl/sharedStrings.xml',v_strings);
    add1file( t_excel, 'xl/_rels/workbook.xml.rels',t_sheet_rels);    
    add1file( t_excel, 'xl/workbook.xml',t_workbook);    
    add1file( t_excel, 'xl/worksheets/_rels/sheet1.xml.rels',v_links);    

    as_zip.finish_zip( t_excel );

    download(p_data      => t_excel,
             p_mime_type => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
             p_file_name => get_file_name (p_app_id,p_page_id,p_region_id)||'.xlsx;'
            );
    dbms_lob.freetemporary(t_excel);
    dbms_lob.freetemporary(v_cells);
    dbms_lob.freetemporary(v_strings);    
    dbms_lob.freetemporary(v_links);    
  END download_excel;

END IR_TO_XLSX;
/
