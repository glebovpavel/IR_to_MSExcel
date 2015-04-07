CREATE OR REPLACE package xml_to_xslx
-- ver 1.0.
IS
  WIDTH_COEFFICIENT constant number := 5;
  
  procedure download_file(p_app_id in number,
                          p_page_id      in number,
                          p_col_length   in varchar2 default null,
                          p_max_rows     in number 
                         ); 
  -- p_col_length is delimetered string COLUMN_NAME=COLUMN_WIDTH,COLUMN_NAME=COLUMN_WIDTH,  etc.
  -- sample: BREAK_ASSIGNED_TO_1=1325,PROJECT=151,TASK_NAME=319,START_DATE=133,
  
  function convert_date_format(p_format in varchar2)
  return varchar2;

  function convert_number_format(p_format in varchar2)
  return varchar2;
  
  function get_max_rows (p_app_id      in number,
                         p_page_id     in number)
  return number;

  
  /*
  -- format test cases
  select xml_to_xslx.convert_date_format('dd.mm.yyyy hh24:mi:ss'),to_char(sysdate,'dd.mm.yyyy hh24:mi:ss') from dual
  union
  select xml_to_xslx.convert_date_format('dd.mm.yyyy hh12:mi:ss'),to_char(sysdate,'dd.mm.yyyy hh12:mi:ss') from dual
  union
  select xml_to_xslx.convert_date_format('day-mon-yyyy'),to_char(sysdate,'day-mon-yyyy') from dual
  union
  select xml_to_xslx.convert_date_format('month'),to_char(sysdate,'month') from dual
  union
  select xml_to_xslx.convert_date_format('RR-MON-DD'),to_char(sysdate,'RR-MON-DD') from dual 
  union
  select xml_to_xslx.convert_number_format('FML999G999G999G999G990D0099'),to_char(123456789/451,'FML999G999G999G999G990D0099') from dual
  union
  select xml_to_xslx.convert_date_format('DD-MON-YYYY HH:MIPM'),to_char(sysdate,'DD-MON-YYYY HH:MIPM') from dual 
  union
  select xml_to_xslx.convert_date_format('fmDay, fmDD fmMonth, YYYY'),to_char(sysdate,'fmDay, fmDD fmMonth, YYYY') from dual 
  */
                    
end;
/


CREATE OR REPLACE PACKAGE body XML_TO_XSLX
is

  STRING_HEIGHT      constant number default 14.4; 

  
  subtype t_color is varchar2(7);
  subtype t_style_string  is varchar2(300);
  subtype t_format_mask  is varchar2(100);
  subtype t_font  is varchar2(50);
  subtype t_large_varchar2  is varchar2(32767);

  BACK_COLOR  constant  t_color default '#C6E0B4';
  
  type t_styles is table of binary_integer index by t_style_string;
  a_styles t_styles;
  type  t_color_list is table of binary_integer index by t_color;
  type  t_font_list is table of binary_integer index by t_font;
  a_font       t_font_list;
  a_back_color t_color_list;
  v_fonts_xml  t_large_varchar2;
  v_back_xml   t_large_varchar2;
  v_styles_xml clob;
  type  t_format_mask_list is table of binary_integer index by t_format_mask;
  a_format_mask_list t_format_mask_list;
  v_format_mask_xml clob;
  
  ------------------------------------------------------------------------------
  t_sheet_rels clob default '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  </Relationships>';
  
  t_workbook clob default '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
  
  t_style_template clob default '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    <cellStyleXfs count="1">
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
    </cellStyleXfs>
    #STYLES#
    <cellStyles count="1">
      <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
    <extLst>
      <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
        <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
      </ext>
    </extLst>
  </styleSheet>';
  
  DEFAULT_FONT constant varchar2(200) := '
   <font>
      <sz val="11" />
      <color theme="1" />
      <name val="Calibri" />
      <family val="2" />
      <scheme val="minor" />
    </font>';

  BOLD_FONT constant varchar2(200) := '
   <font>
      <b />
      <sz val="11" />
      <color theme="1" />
      <name val="Calibri" />
      <family val="2" />
      <scheme val="minor" />
    </font>';

  FONTS_CNT constant  binary_integer := 2;   
  
  
  DEFAULT_FILL constant varchar2(200) :=  '
    <fill>
      <patternFill patternType="none" />
    </fill>
    <fill>
      <patternFill patternType="gray125" />
    </fill>'; 
  DEFAULT_FILL_CNT constant  binary_integer := 2; 
  
  DEFAULT_STYLE constant varchar2(200) := '
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';

  AGGREGATE_STYLE constant varchar2(250) := '
      <xf numFmtId="#FMTID#" borderId="1" fillId="0" fontId="1" xfId="0" applyAlignment="1" applyFont="1" applyBorder="1">
         <alignment wrapText="1" horizontal="right"  vertical="top"/>
     </xf>';
  
     
  HEADER_STYLE constant varchar2(200) := '     
      <xf numFmtId="0" borderId="0" #FILL# fontId="1" xfId="0" applyAlignment="1" applyFont="1" >
         <alignment wrapText="0" horizontal="#ALIGN#"/>
     </xf>';  
  
  
  DEFAULT_STYLES_CNT constant  binary_integer := 1;     

  FORMAT_MASK_START_WITH  constant  binary_integer := 164;  
  
  FORMAT_MASK constant varchar2(100) := '
      <numFmt numFmtId="'||FORMAT_MASK_START_WITH||'" formatCode="dd.mm.yyyy"/>';
  
  FORMAT_MASK_CNT constant binary_integer default 1; 
  
  
  cursor cur_row(p_xml xmltype) is 
  SELECT rownum coll_num,
                                 extractvalue(COLUMN_VALUE, 'CELL/@background-color') AS background_color,
                                 extractvalue(COLUMN_VALUE, 'CELL/@color') AS font_color, 
                                 extractvalue(COLUMN_VALUE, 'CELL/@data-type') AS data_type,
                                 extractvalue(COLUMN_VALUE, 'CELL/@value') AS cell_value,
                                 extractvalue(column_value, 'CELL') as cell_text
                          from table (select xmlsequence(extract(p_xml,'DOCUMENT/DATA/ROWSET/ROW/CELL')) from dual);
  ------------------------------------------------------------------------------
  function convert_date_format(p_format in varchar2)
  return varchar2
  is
    v_str      varchar2(100);
  begin
    v_str := p_format;
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
    
    v_str := replace(v_str,'HH24','h');
    v_str := regexp_replace(v_str,'(\W)','\\\1');
    v_str := regexp_replace(v_str,'HH12([^ ]+)','h\1 AM/PM');    
    v_str := replace(v_str,'AM\/PM',' AM/PM');
    
    
    return v_str;
  end convert_date_format;
  ------------------------------------------------------------------------------
  function convert_number_format(p_format in varchar2)
  return varchar2
  is
    v_str      varchar2(100);
  begin
    v_str := p_format;
    v_str := upper(v_str);
    --number
    v_str := replace(v_str,'9','#');
    v_str := replace(v_str,'D','.');
    v_str := replace(v_str,'G',',');
    v_str := replace(v_str,'FM','');
    --currency
    v_str := replace(v_str,'L',convert('&quot;'||rtrim(to_char(0,'FML0'),'0')||'&quot;','UTF8'));    
    
    v_str := regexp_substr(v_str,'.G?[^G]+$');
    return v_str;
  end convert_number_format;  
  ------------------------------------------------------------------------------
  function add_font(p_font in t_color,p_bold in varchar2 default null)
  return binary_integer 
  is  
   v_index t_font;
  begin
    v_index := p_font||p_bold;
    
    if not a_font.exists(v_index) then
      a_font(v_index) := a_font.count + 1;
      v_fonts_xml := v_fonts_xml||
        '<font>'||chr(10)||
        '   <sz val="11" />'||chr(10)||
        '   <color rgb="'||ltrim(p_font,'#')||'" />'||chr(10)||
        '   <name val="Calibri" />'||chr(10)||
        '   <family val="2" />'||chr(10)||
        '   <scheme val="minor" />'||
        '</font>'||chr(10);
        
        return a_font.count + FONTS_CNT - 1; --start with 0
     else
       return a_font(v_index) + FONTS_CNT - 1;
     end if;
  end add_font;
  ------------------------------------------------------------------------------
  function  add_back_color(p_back in t_color)
  return binary_integer 
  is
  begin
    if not a_back_color.exists(p_back) then
      a_back_color(p_back) := a_back_color.count + 1;
      v_back_xml := v_back_xml||
        '<fill>'||chr(10)||
        '   <patternFill patternType="solid">'||chr(10)||
        '     <fgColor rgb="'||ltrim(p_back,'#')||'" />'||chr(10)||
        '     <bgColor indexed="64" />'||chr(10)||
        '   </patternFill>'||
        '</fill>'||chr(10);
        
        return  a_back_color.count + DEFAULT_FILL_CNT - 1; --start with 0
     else
       return a_back_color(p_back) + DEFAULT_FILL_CNT - 1;
     end if;
  end add_back_color;
  ------------------------------------------------------------------------------
  function  add_format_mask(p_mask in varchar2)
  return binary_integer 
  is
  begin
    if not a_format_mask_list.exists(p_mask) then
      a_format_mask_list(p_mask) := a_format_mask_list.count + 1;
      v_format_mask_xml := v_format_mask_xml||
        '<numFmt numFmtId="'||(FORMAT_MASK_START_WITH + a_format_mask_list.count)||'" formatCode="'||p_mask||'"/>'||chr(10);

        return  a_format_mask_list.count + FORMAT_MASK_CNT + FORMAT_MASK_START_WITH - 1; 
     else
       return a_format_mask_list(p_mask) + FORMAT_MASK_CNT + FORMAT_MASK_START_WITH - 1;
     end if;
  end add_format_mask;  
  ------------------------------------------------------------------------------
  function get_font_colors_xml 
  return clob
  is
  begin
    return to_clob('<fonts count="'||(a_font.count + FONTS_CNT)||'" x14ac:knownFonts="1">'||
                   DEFAULT_FONT||chr(10)||
                   BOLD_FONT||chr(10)||
                   v_fonts_xml||chr(10)||
                   '</fonts>'||chr(10));
  end get_font_colors_xml;  
  ------------------------------------------------------------------------------
  function get_back_colors_xml 
  return clob
  is
  begin
    return to_clob('<fills count="'||(a_back_color.count + DEFAULT_FILL_CNT)||'">'||
                   DEFAULT_FILL||
                   v_back_xml||chr(10)||
                   '</fills>'||chr(10));
  end get_back_colors_xml;  
  ------------------------------------------------------------------------------  
  function get_format_mask_xml 
  return clob
  is
  begin
    return to_clob('<numFmts count="'||(a_format_mask_list.count + FORMAT_MASK_CNT)||'">'||
                   FORMAT_MASK||
                   v_format_mask_xml||chr(10)||
                   '</numFmts>'||chr(10));
  end get_format_mask_xml;  
  ------------------------------------------------------------------------------  
  function get_cellXfs_xml
  return clob
  is
  begin
    return to_clob('<cellXfs count="'||(a_styles.count + DEFAULT_STYLES_CNT)||'">'||
                   DEFAULT_STYLE||chr(10)||
                   v_styles_xml||chr(10)||
                   '</cellXfs>'||chr(10));
  end get_cellXfs_xml;
  ------------------------------------------------------------------------------  
  function get_num_fmt_id(p_data_type   in varchar2,
                          p_format_mask in varchar2)
  return binary_integer
  is
  begin
      if p_data_type = 'NUMBER' then 
       if p_format_mask is null then
          return 0;
        else
          return add_format_mask(convert_number_format(p_format_mask)); 
        end if;
      elsif p_data_type = 'DATE' then 
        if p_format_mask is null then
          return FORMAT_MASK_START_WITH;  -- default date format
        else
          return add_format_mask(convert_date_format(p_format_mask)); 
        end if;
      else
        return 2;  -- default  string format
      end if;  
  
  end get_num_fmt_id;  
  ------------------------------------------------------------------------------  
  -- get style_id for existent styles or 
  -- add new style and return style_id
  function get_style_id(p_font        in t_color,
                        p_back        in t_color,
                        p_data_type   in varchar2,
                        p_format_mask in varchar2,
                        p_align       in varchar2)
  return binary_integer
  is
    v_style           t_style_string;
    v_font_color_id   binary_integer;
    v_back_color_id   binary_integer;
    v_style_xml       t_large_varchar2;
  begin
    v_style := nvl(p_font,'      ')||nvl(p_back,'       ')||p_data_type||p_format_mask||p_align;
    --   
    if a_styles.exists(v_style) then
      return a_styles(v_style)   - 1 + DEFAULT_STYLES_CNT;
    else
      a_styles(v_style) := a_styles.count + 1;
      
      v_style_xml := '<xf borderId="0" xfId="0" ';

      v_style_xml := v_style_xml||replace(' numFmtId="#FMTID#" ','#FMTID#',get_num_fmt_id(p_data_type,p_format_mask));
      
      if p_font is not null then
        v_font_color_id := add_font(p_font);
        v_style_xml := v_style_xml||' fontId="'||v_font_color_id||'"  applyFont="1" ';
      else
        v_style_xml := v_style_xml||' fontId="0" '; --default font
      end if;
      
      if p_back is not null then
        v_back_color_id := add_back_color(p_back);
        v_style_xml := v_style_xml||' fillId="'||v_back_color_id||'"  applyFill="1" ';
      else
        v_style_xml := v_style_xml||' fillId="0" '; --default fill 
      end if;
      
      v_style_xml := v_style_xml||'>'||chr(10);
      v_style_xml := v_style_xml||'<alignment wrapText="1"';
      if p_align is not null then
        v_style_xml := v_style_xml||' horizontal="'||lower(p_align)||'" ';
      end if;
      v_style_xml := v_style_xml||'/>'||chr(10);
      
      v_style_xml := v_style_xml||'</xf>'||chr(10);      
      v_styles_xml := v_styles_xml||to_clob(v_style_xml);      
      return a_styles.count  - 1 + DEFAULT_STYLES_CNT;
    end if;  
  
  end get_style_id;
  ------------------------------------------------------------------------------  
  -- get style_id for existent styles or 
  -- add new style and return style_id
  function get_aggregate_style_id(p_font        in t_color,                               
                                  p_back        in t_color,
                                  p_data_type   in varchar2,
                                  p_format_mask in varchar2)
  return binary_integer
  is
    v_style           t_style_string;
    v_style_xml       t_large_varchar2;
    v_num_fmt_id        binary_integer;
  begin
    v_style := nvl(p_font,'      ')||nvl(p_back,'       ')||p_data_type||p_format_mask||'AGGREGATE';
    --   
    if a_styles.exists(v_style) then
      return a_styles(v_style)  - 1 + DEFAULT_STYLES_CNT;
    else
      a_styles(v_style) := a_styles.count + 1;

      v_style_xml  := replace(AGGREGATE_STYLE,'#FMTID#',get_num_fmt_id(p_data_type,p_format_mask)) ||chr(10);      
      v_styles_xml := v_styles_xml||v_style_xml;      
      return a_styles.count  - 1 + DEFAULT_STYLES_CNT;
    end if;  
  
  end get_aggregate_style_id;  
  ------------------------------------------------------------------------------
  function get_header_style_id(p_back        in t_color default BACK_COLOR,
                               p_align       in varchar2,
                               p_border      in boolean default false)
  return binary_integer
  is
    v_style             t_style_string;
    v_style_xml         t_large_varchar2;
    v_num_fmt_id        binary_integer;
    v_back_color_id     binary_integer;
  begin
    v_style := '      '||nvl(p_back,'       ')||'CHARHEADER'||p_align;
    --   
    if a_styles.exists(v_style) then
      return a_styles(v_style)  - 1 + DEFAULT_STYLES_CNT;
    else
      a_styles(v_style) := a_styles.count + 1;
  
      if p_back is not null then
        v_back_color_id := add_back_color(p_back);
        v_style_xml := replace(HEADER_STYLE,'#FILL#',' fillId="'||v_back_color_id||'"  applyFill="1" ' );
      else
        v_style_xml := replace(HEADER_STYLE,'#FILL#',' fillId="0" '); --default fill 
      end if;      

      v_style_xml  := replace(v_style_xml,'#ALIGN#',lower(p_align))||chr(10);
      v_styles_xml := v_styles_xml||v_style_xml;      
      return a_styles.count  - 1 + DEFAULT_STYLES_CNT;
    end if;  
  
  end get_header_style_id;    
  ------------------------------------------------------------------------------
  
  function  get_styles_xml
  return clob
  is
    v_template clob;
  begin
     v_template := replace(t_style_template,'#FORMAT_MASKS#',get_format_mask_xml);     
     v_template := replace(v_template,'#FONTS#',get_font_colors_xml);
     v_template := replace(v_template,'#FILLS#',get_back_colors_xml);
     v_template := replace(v_template,'#STYLES#',get_cellXfs_xml);
     
     return v_template;
  end get_styles_xml;
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
  PROCEDURE add( p_clob IN OUT NOCOPY CLOB
               , p_vc_buffer IN OUT NOCOPY VARCHAR2
               , p_vc_addition IN VARCHAR2
               , p_eof IN BOOLEAN DEFAULT FALSE
               )
  AS
  BEGIN
     
    -- Standard Flow
    IF NVL(LENGTHB(p_vc_buffer), 0) + NVL(LENGTHB(p_vc_addition), 0) < 32767 THEN
      p_vc_buffer := p_vc_buffer || p_vc_addition;
    ELSE
      IF p_clob IS NULL THEN
        dbms_lob.createtemporary(p_clob, TRUE);
      END IF;
      dbms_lob.writeappend(p_clob, length(p_vc_buffer), p_vc_buffer);
      p_vc_buffer := p_vc_addition;
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
  FUNCTION get_cell_name(p_coll IN binary_integer,
                         p_row  IN binary_integer)
  RETURN varchar2
  IS   
  BEGIN
   if p_coll > 26 then
     return chr(64 + trunc(p_coll/26))||chr(64 + p_coll - trunc(p_coll/26)*26 +1)||p_row;
   ELSE
     return chr( 64 + p_coll)||p_row;
   end if;  
  end get_cell_name;
  ------------------------------------------------------------------------------  
  function get_colls_width_xml(p_width_str    in varchar2,
                               p_xml          in xmltype,
                               p_coefficient  in number)
  return clob
  is
    subtype t_column_alias is varchar2(255);
    v_xml                  clob;    
    a_col_name_plus_width  apex_application_global.vc_arr2;    
    v_col_alias            t_column_alias;
    v_col_width            number;
    type   t_width_list    is table of integer index by t_column_alias;
    a_width_list           t_width_list;
    v_coefficient          number;
    v_is_custom            boolean default false;
  begin      
    a_col_name_plus_width := APEX_UTIL.STRING_TO_TABLE(p_width_str,','); 
    
    v_coefficient := nvl(p_coefficient,WIDTH_COEFFICIENT);
    if v_coefficient =  0 then
      v_coefficient := WIDTH_COEFFICIENT;
    end if;     
    
    -- init associative array by colunn width
    for i in 1..a_col_name_plus_width.count loop      
      
      if a_col_name_plus_width(i) is not null then
        begin                    
          v_col_alias  := regexp_replace(a_col_name_plus_width(i),'\=\d+,?$','');
          v_col_width  := ltrim(rtrim(regexp_substr(a_col_name_plus_width(i),'\=\d+,?$'),','),'=');          
          if v_col_width is not null and v_col_alias is not null then            
            a_width_list(v_col_alias) := round(to_number(v_col_width) / v_coefficient);            
          end if;
         exception
           when others then             
             null;  -- this functionality is not important
         end ;
      end if;
    end loop;
    --set column width
    v_xml:= v_xml||to_clob('<cols>'||chr(10));
    for i in (select  rownum rn,
                      extractvalue(column_value, 'CELL') as column_header,
                      extractvalue(COLUMN_VALUE, 'CELL/@column-alias') AS column_alias,  
                      extractvalue(COLUMN_VALUE, 'CELL/@data-type') AS data_type
               from table (select xmlsequence(extract(p_xml,'DOCUMENT/DATA/HEADER/CELL')) from dual))
    loop                    
      if a_width_list.exists(i.column_alias)  then        
        v_xml:= v_xml||to_clob('<col min="'||i.rn||'" max="'||i.rn||'" width="'||a_width_list(i.column_alias)||'" customWidth="1" />'||chr(10));        
        v_is_custom := true;        
      end if;  
    end loop;
    v_xml:= v_xml||to_clob('</cols>'||chr(10));       
    
    if v_is_custom then
      return v_xml;      
    else
      return to_clob('');
    end if; 
  exception
    when others then 
      raise_application_error(-20001,'get_colls_width_xml: '||SQLERRM);
  end get_colls_width_xml;
 
  ------------------------------------------------------------------------------  
  procedure add1file
    ( p_zipped_blob in out blob
    , p_name in varchar2
    , p_content in clob
    )
  is
   v_desc_offset pls_integer := 1;
   v_src_offset  PLS_INTEGER := 1;
   v_lang        pls_integer := 0;
   v_warning     pls_integer := 0;
   v_blob        BLOB;
  begin
    dbms_lob.createtemporary(v_blob,true);    
    dbms_lob.converttoblob(v_blob,p_content, dbms_lob.getlength(p_content), v_desc_offset, v_src_offset, dbms_lob.default_csid, v_lang, v_warning);
    as_zip.add1file( p_zipped_blob, p_name, v_blob);
    dbms_lob.freetemporary(v_blob);
  end add1file;  
  ------------------------------------------------------------------------------
  procedure get_excel(p_xml in xmltype,
                      v_clob in out nocopy clob,
                      v_strings_clob in out nocopy clob,
                      p_width_str in varchar2,
                      p_coefficient  in number)
  IS
    v_strings          apex_application_global.vc_arr2;
    v_rownum           binary_integer default 1;
    v_colls_count      binary_integer default 0;
    v_agg_clob         clob;
    v_agg_strings_cnt  binary_integer default 1; 
    v_style_id         binary_integer; 
    v_buffer           t_large_varchar2;
    v_agg_buffer       t_large_varchar2; 
    v_string_buffer    t_large_varchar2;
    --
    procedure print_char_cell(p_coll      in binary_integer, 
                              p_row       in binary_integer, 
                              p_string    in varchar2,
                              p_clob      in out nocopy CLOB,
                              p_buffer    IN OUT NOCOPY VARCHAR2,
                              p_style_id  in number default null
                             )
    is
     v_style varchar2(20);
    begin
      if p_style_id is not null then
       v_style := ' s="'||p_style_id||'" ';
      end if;
      add(p_clob,p_buffer,'<c r="'||get_cell_name(p_coll,p_row)||'" t="s" '||v_style||'>'||chr(10)
                         ||'<v>' || to_char(v_strings.count)|| '</v>'||chr(10)                 
                         ||'</c>'||chr(10));
      v_strings(v_strings.count + 1) := p_string;
    end print_char_cell;
    --
    procedure print_number_cell(p_coll      in binary_integer, 
                                p_row       in binary_integer, 
                                p_value     in varchar2,
                                p_clob      in out nocopy clob,
                                p_buffer    IN OUT NOCOPY VARCHAR2,
                                p_style_id  in number default null
                               )
    is
      v_style varchar2(20);
    begin
      if p_style_id is not null then
       v_style := ' s="'||p_style_id||'" ';
      end if;
      add(p_clob,p_buffer,'<c r="'||get_cell_name(p_coll,p_row)||'" '||v_style||'>'||chr(10)
                         ||'<v>'||p_value|| '</v>'||chr(10)
                         ||'</c>'||chr(10));
    
    end print_number_cell;
    --
  begin
     pragma inline(add,'YES');     
     pragma inline(get_cell_name,'YES');
     
     dbms_lob.createtemporary(v_agg_clob,true);
     
     --!
     add(v_clob,v_buffer,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'||chr(10));
     add(v_clob,v_buffer,'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <dimension ref="A1"/>
        <sheetViews>
          <sheetView tabSelected="1" workbookViewId="0">
            <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
            <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
           </sheetView>
          </sheetViews>
        <sheetFormatPr baseColWidth="10" defaultColWidth="10" defaultRowHeight="15"/>'||chr(10),TRUE);

     v_clob := v_clob||get_colls_width_xml(p_width_str,p_xml,p_coefficient);
     --!
     add(v_clob,v_buffer,'<sheetData>'||chr(10));

     --column header
     add(v_clob,v_buffer,'<row>'||chr(10));     
     for i in (select  extractvalue(column_value, 'CELL') as column_header,
                       extractvalue(COLUMN_VALUE, 'CELL/@header_align') AS align
               from table (select xmlsequence(extract(p_xml,'DOCUMENT/DATA/HEADER/CELL')) from dual))
     loop          
       v_colls_count := v_colls_count + 1;
       print_char_cell( p_coll      => v_colls_count, 
                        p_row       => v_rownum, 
                        p_string    => i.column_header,
                        p_clob      => v_clob,
                        p_buffer    => v_buffer,
                        p_style_id  => get_header_style_id(p_align  => i.align)
                      );
     end loop; 
     v_rownum := v_rownum + 1;
     add(v_clob,v_buffer,'</row>'||chr(10));     
     
     <<rowset>>     
     for rowset_xml in (select column_value as rowset,
                               extractvalue(COLUMN_VALUE, 'ROWSET/BREAK_HEADER') AS break_header
                        from table (select xmlsequence(extract(p_xml,'DOCUMENT/DATA/ROWSET')) from dual)
                       ) 
     loop
       --break header
       if rowset_xml.break_header is not null then
         add(v_clob,v_buffer,'<row>'||chr(10));
         print_char_cell( p_coll      => 1,
                          p_row       => v_rownum,
                          p_string    => rowset_xml.break_header,
                          p_clob      => v_clob,
                          p_buffer    => v_buffer,
                          p_style_id  => get_header_style_id(p_back => NULL,p_align  => 'left')
                         );         
         v_rownum := v_rownum + 1;
         add(v_clob,v_buffer,'</row>'||chr(10));
       end if;
       
       <<cells>>
       for row_xml in (select column_value as row_ from table (select xmlsequence(extract(rowset_xml.rowset,'ROWSET/ROW')) from dual)) loop
         add(v_clob,v_buffer,'<row>'||chr(10));
         FOR cell_xml IN (SELECT rownum coll_num,
                                 extractvalue(COLUMN_VALUE, 'CELL/@background-color') AS background_color,
                                 extractvalue(COLUMN_VALUE, 'CELL/@color') AS font_color, 
                                 extractvalue(COLUMN_VALUE, 'CELL/@data-type') AS data_type,
                                 extractvalue(COLUMN_VALUE, 'CELL/@value') AS cell_value,
                                 extractvalue(COLUMN_VALUE, 'CELL/@format_mask') AS format_mask,
                                 extractvalue(COLUMN_VALUE, 'CELL/@align') AS align,
                                 extractvalue(COLUMN_VALUE, 'CELL') AS cell_text
                          from table (select xmlsequence(extract(row_xml.row_,'ROW/CELL')) from dual))
          loop
            begin            
              v_style_id := get_style_id(p_font         => cell_xml.font_color,
                                         p_back         => cell_xml.background_color,
                                         p_data_type    => cell_xml.data_type,
                                         p_format_mask  => cell_xml.format_mask,
                                         p_align        => cell_xml.align
                                        );
              if cell_xml.data_type in ('NUMBER') then
                  print_number_cell(p_coll      => cell_xml.coll_num, 
                                    p_row       => v_rownum, 
                                    p_value     => cell_xml.cell_value,
                                    p_clob      => v_clob,
                                    p_buffer    => v_buffer,
                                    p_style_id  => v_style_id
                                   );
                  
              elsif cell_xml.data_type in ('DATE') then
                 add(v_clob,v_buffer,'<c r="'||get_cell_name(cell_xml.coll_num,v_rownum)||'"  s="'||v_style_id||'">'||chr(10)
                                    ||'<v>'||cell_xml.cell_value|| '</v>'||chr(10)
                                    ||'</c>'||chr(10)
                     );
              else --STRING
                  print_char_cell(p_coll        => cell_xml.coll_num,
                                  p_row         => v_rownum,
                                  p_string      => cell_xml.cell_text,
                                  p_clob        => v_clob,
                                  p_buffer      => v_buffer,
                                  p_style_id    => v_style_id
                                  );
              END IF;              
            exception
              WHEN no_data_found THEN
                null;
            end;
         end loop;         
         add(v_clob,v_buffer,'</row>'||chr(10));         
         v_rownum := v_rownum + 1;
       end loop cells;
       
       DBMS_LOB.TRIM(v_agg_clob,0);
       v_agg_buffer := '';       
       v_agg_strings_cnt := 1;       
       <<aggregates>>       
       for row_xml in (select column_value as row_ from table (select xmlsequence(extract(rowset_xml.rowset,'ROWSET/AGGREGATE')) from dual)) loop
         for cell_xml_agg in (select rownum coll_num,
                                 extractvalue(COLUMN_VALUE, 'CELL') AS cell_text,
                                 extractvalue(COLUMN_VALUE, 'CELL/@value') AS cell_value,
                                 extractvalue(COLUMN_VALUE, 'CELL/@format_mask') AS format_mask
                          from table (select xmlsequence(extract(row_xml.row_,'AGGREGATE/CELL')) from dual))
         loop           
           v_agg_strings_cnt := greatest(length(regexp_replace('[^:]','')) + 1,v_agg_strings_cnt);
             v_style_id := get_aggregate_style_id(p_font         => '',
                                        p_back         => '',
                                        p_data_type    => 'CHAR',
                                        p_format_mask  => '');
             print_char_cell(p_coll       => cell_xml_agg.coll_num,
                             p_row        => v_rownum,
                             p_string     => rtrim(cell_xml_agg.cell_text,chr(10)),
                             p_clob       => v_agg_clob,
                             p_buffer    =>  v_agg_buffer,           
                             p_style_id   => v_style_id);
         end loop;
         add(v_clob,v_buffer,'<row ht="'||(v_agg_strings_cnt * STRING_HEIGHT)||'">'||chr(10),TRUE);
         add(v_agg_clob,v_agg_buffer,' ',TRUE);
         dbms_lob.copy( dest_lob => v_clob,
                        src_lob => v_agg_clob,
                        amount => dbms_lob.getlength(v_agg_clob),
                        dest_offset => dbms_lob.getlength(v_clob),
                        src_offset => 1);
         add(v_clob,v_buffer,'</row>'||chr(10));
         v_rownum := v_rownum + 1;
       end loop aggregates;   
     end loop rowset;     
     
     add(v_clob,v_buffer,'</sheetData>'||chr(10));
     --if p_autofilter then
     add(v_clob,v_buffer,'<autoFilter ref="A1:' || get_cell_name(v_colls_count,v_rownum) || '"/>');
     --end if;
     add(v_clob,v_buffer,'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'||chr(10),TRUE);
     
     add(v_strings_clob,v_string_buffer,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'||chr(10));
     add(v_strings_clob,v_string_buffer,'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' || v_strings.count() || '" uniqueCount="' || v_strings.count() || '">'||chr(10));
        
     for i in 1 .. v_strings.count() loop
        add(v_strings_clob,v_string_buffer,'<si><t>'||dbms_xmlgen.convert( substr( v_strings( i ), 1, 32000 ) ) || '</t></si>'||chr(10));
     end loop; 
     add(v_strings_clob,v_string_buffer,'</sst>'||chr(10),TRUE);
     
     dbms_lob.freetemporary(v_agg_clob);
  end get_excel;
  ------------------------------------------------------------------------------
  function get_max_rows (p_app_id      in number,
                         p_page_id     in number)
  return number
  is 
    v_max_row_count number;
  begin
    select max_row_count 
    into v_max_row_count
    from APEX_APPLICATION_PAGE_IR
    where application_id = p_app_id
      and page_id = p_page_id;
       
     return v_max_row_count;
  end get_max_rows;   
  ------------------------------------------------------------------------------  
  function get_file_name (p_app_id      in number,
                          p_page_id     in number)
  return varchar2
  is 
    v_filename varchar2(255);
  begin
    select filename 
    into v_filename
    from APEX_APPLICATION_PAGE_IR
    where application_id = p_app_id
      and page_id = p_page_id;
       
     return apex_plugin_util.replace_substitutions(nvl(v_filename,'Excel'));
  end get_file_name;   
  ------------------------------------------------------------------------------
  procedure download_file(p_app_id      in number,
                          p_page_id     in number,
                          p_col_length   in varchar2 default null,
                          p_max_rows     in number 
                         )
  is
    t_template blob;
    t_excel    blob;
    v_cells    clob;
    v_strings  clob;
    v_xml_data xmltype;
    zip_files  as_zip.file_list;
  begin    
    pragma inline(get_excel,'YES');     
    dbms_lob.createtemporary(t_excel,true);    
    dbms_lob.createtemporary(v_cells,true);
    dbms_lob.createtemporary(v_strings,true);
    
    
    select file_content
    into t_template
    from apex_appl_plugin_files 
    where file_name = 'ExcelTemplate.zip'
      and application_id = p_app_id;
    
    zip_files  := as_zip.get_file_list( t_template );
    for i in zip_files.first() .. zip_files.last loop
      as_zip.add1file( t_excel, zip_files( i ), as_zip.get_file( t_template, zip_files( i ) ) );
    end loop;
    
    v_xml_data := IR_TO_XML.get_report_xml(p_app_id => p_app_id,
                          p_page_id => p_page_id,                                
                          p_get_page_items => 'N',
                          p_items_list  => null,
                          p_max_rows  => p_max_rows
                         );
    
    
    get_excel(v_xml_data,v_cells,v_strings,p_col_length,WIDTH_COEFFICIENT);
    add1file( t_excel, 'xl/styles.xml', get_styles_xml);
    add1file( t_excel, 'xl/worksheets/Sheet1.xml', v_cells);
    add1file( t_excel, 'xl/sharedStrings.xml',v_strings);
    add1file( t_excel, 'xl/_rels/workbook.xml.rels',t_sheet_rels);    
    add1file( t_excel, 'xl/workbook.xml',t_workbook);    
    
    as_zip.finish_zip( t_excel );
      
    htp.flush;
    htp.init();
    owa_util.mime_header( wwv_flow_utilities.get_excel_mime_type, false );
    htp.print( 'Content-Length: ' || dbms_lob.getlength( t_excel ) );
    htp.print( 'Content-disposition: attachment; filename='||get_file_name (p_app_id,p_page_id)||'.xlsx;' );
    owa_util.http_header_close;
    wpg_docload.download_file( t_excel );
    dbms_lob.freetemporary(t_excel);
    dbms_lob.freetemporary(v_cells);
    dbms_lob.freetemporary(v_strings);    
  end download_file;
  
end;
/
