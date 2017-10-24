create or replace PACKAGE  IR_TO_XML 
  AUTHID CURRENT_USER
as    

  PROCEDURE get_report_xml(p_app_id          IN NUMBER,
                           p_page_id         IN NUMBER,     
                           p_region_id       IN NUMBER,
                           p_return_type     IN CHAR DEFAULT 'X', -- "Q" for debug information "X" for XML-Data
                           p_get_page_items  IN CHAR DEFAULT 'N', -- Y,N - include page items in XML
                           p_items_list      IN VARCHAR2,         -- "," delimetered list of items that for including in XML
                           p_collection_name IN VARCHAR2,         -- name of APEX COLLECTION to save XML, when null - download as file
                           p_max_rows        IN NUMBER            -- maximum rows for export                            
                          );
  
  --return debug information
  function get_log return clob;
  
  -- get XML 
  function get_report_xml(p_app_id          IN NUMBER,
                          p_page_id         IN NUMBER,  
                          p_region_id       IN NUMBER,
                          p_get_page_items  IN CHAR DEFAULT 'N', -- Y,N - include page items in XML
                          p_items_list      IN VARCHAR2,         -- "," delimetered list of items that for including in XML
                          p_max_rows        IN NUMBER            -- maximum rows for export                            
                         )
  return xmltype;     
  /* 
    function to handle cases of 'in' and 'not in' conditions for highlights
   	used in cursor cur_highlight
    
    Author: Srihari Ravva
  */ 
  function get_highlight_in_cond_sql(p_condition_expression  in APEX_APPLICATION_PAGE_IR_COND.CONDITION_EXPRESSION%TYPE,
                                     p_condition_sql         in APEX_APPLICATION_PAGE_IR_COND.CONDITION_SQL%TYPE,
                                     p_condition_column_name in APEX_APPLICATION_PAGE_IR_COND.CONDITION_COLUMN_NAME%TYPE)
  return varchar2; 
                              
END IR_TO_XML;
/
create or replace PACKAGE BODY IR_TO_XML   
as    
/*
** Minor bugfixes by J.P.Lourens  9-Oct-2016
*/  
  subtype largevarchar2 is varchar2(32767); 
  subtype columntype is varchar2(15); 
  subtype formatmask is varchar2(100);  

  format_error EXCEPTION;
  PRAGMA EXCEPTION_INIT(format_error, -01830);
  date_format_error EXCEPTION;
  PRAGMA EXCEPTION_INIT(format_error, -01821);
  conversion_error exception;
  PRAGMA EXCEPTION_INIT(conversion_error,-06502);
  
  type t_date_table is table of date index by binary_integer;  
  type t_number_table is table of number index by binary_integer;  
  
  cursor cur_highlight(p_report_id in APEX_APPLICATION_PAGE_IR_RPT.REPORT_ID%TYPE,
                       p_delimetered_column_list in varchar2) 
  IS
  select rez.* ,
       rownum COND_NUMBER,
       'HIGHLIGHT_'||rownum COND_NAME
  from (
  select report_id,
         case when condition_operator in ('not in', 'in') then
		         IR_TO_XML.get_highlight_in_cond_sql(CONDITION_EXPRESSION,CONDITION_SQL,CONDITION_COLUMN_NAME)
		      else 
		         replace(replace(replace(replace(condition_sql,'#APXWS_EXPR#',''''||CONDITION_EXPRESSION||''''),'#APXWS_EXPR2#',''''||CONDITION_EXPRESSION2||''''),'#APXWS_HL_ID#','1'),'#APXWS_CC_EXPR#','"'||CONDITION_COLUMN_NAME||'"') 
		     end condition_sql,
         CONDITION_COLUMN_NAME,
         CONDITION_ENABLED,
         HIGHLIGHT_ROW_COLOR,
         HIGHLIGHT_ROW_FONT_COLOR,
         HIGHLIGHT_CELL_COLOR,
         HIGHLIGHT_CELL_FONT_COLOR      
    from APEX_APPLICATION_PAGE_IR_COND
    where condition_type = 'Highlight'
      and report_id = p_report_id
      and instr(':'||p_delimetered_column_list||':',':'||CONDITION_COLUMN_NAME||':') > 0
      and condition_enabled = 'Yes'
      order by --rows highlights first 
             nvl2(HIGHLIGHT_ROW_COLOR,1,0) desc, 
             nvl2(HIGHLIGHT_ROW_FONT_COLOR,1,0) desc,
             HIGHLIGHT_SEQUENCE 
    ) rez;
  
  type t_col_names is table of apex_application_page_ir_col.report_label%type index by apex_application_page_ir_col.column_alias%type;
  type t_col_format_mask is table of APEX_APPLICATION_PAGE_IR_COMP.computation_format_mask%TYPE index by APEX_APPLICATION_PAGE_IR_COL.column_alias%TYPE;
  type t_header_alignment is table of APEX_APPLICATION_PAGE_IR_COL.heading_alignment%TYPE index by APEX_APPLICATION_PAGE_IR_COL.column_alias%TYPE;
  type t_column_alignment is table of apex_application_page_ir_col.column_alignment%type index by apex_application_page_ir_col.column_alias%type;
  type t_column_types is table of apex_application_page_ir_col.column_type%type index by binary_integer;
  type t_highlight is table of cur_highlight%ROWTYPE index by binary_integer;
  
  type ir_report is record
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
    skipped_columns           BINARY_INTEGER default 0, -- when scpecial coluns like apxws_row_pk is used
    start_with                BINARY_INTEGER default 0, -- position of first displayed column in query
    end_with                  BINARY_INTEGER default 0, -- position of last displayed column in query    
    agg_cols_cnt              BINARY_INTEGER default 0, 
    hidden_cols_cnt           BINARY_INTEGER default 0, 
    column_names              t_col_names,       -- column names in report header
    col_format_mask           t_col_format_mask, -- format like $3849,56
    row_highlight             t_highlight,
    col_highlight             t_highlight,
    header_alignment          t_header_alignment,
    column_alignment          t_column_alignment,
    column_types              t_column_types  
   );  
   
   TYPE t_cell_data IS record
   (
     VALUE           VARCHAR2(100),
     text            largevarchar2,
     datatype        VARCHAR2(50)
   );  
  l_report                   ir_report;   
  v_debug                    clob;  
  v_debug_buffer             largevarchar2; 
  
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
  procedure log(p_message in varchar2,p_eof IN BOOLEAN DEFAULT FALSE)
  is
  begin
  /* logigging  ffrffe */
    add(v_debug,v_debug_buffer,p_message||chr(10),p_eof);
    apex_debug_message.log_message(p_message => substr(p_message,1,32767),
                                   p_enabled => false,
                                   p_level   => 4);
  end log; 
  ------------------------------------------------------------------------------
  function get_log
  return clob
  is
  begin
    log('LogFinish',TRUE);
    return v_debug;
  end  get_log; 
  ------------------------------------------------------------------------------
  
  function bcoll(p_font_color    in varchar2 default null,
                 p_back_color    in varchar2 default null,
                 p_align         in varchar2 default null,
                 p_width         in varchar2 default null,
                 p_column_alias  IN VARCHAR2 DEFAULT NULL,
                 p_colmn_type    IN VARCHAR2 DEFAULT NULL,
                 p_value         IN VARCHAR2 DEFAULT NULL,
                 p_format_mask   IN VARCHAR2 DEFAULT NULL,
                 p_header_align  in varchar2 default null) 
  return varchar2
  is
    v_str varchar2(500);
  begin
    v_str := v_str||'<CELL ';
    if p_column_alias is not null then v_str := v_str||'column-alias="'||APEX_ESCAPE.HTML_ATTRIBUTE(p_column_alias)||'" '; end if; 
    if p_font_color is not null then v_str := v_str||'color="'||p_font_color||'" '; end if;
    if p_colmn_type is not null then V_STR := V_STR||'data-type="'||p_colmn_type||'" '; end if;
    if p_back_color is not null then v_str := v_str||'background-color="'||p_back_color||'" '; end if;
    if p_align is not null then V_STR := V_STR||'align="'||lower(p_align)||'" '; end if;
    IF p_width IS NOT NULL THEN v_str := v_str||'width="'||p_width||'" '; END IF;        
    IF p_value IS NOT NULL THEN v_str := v_str||'value="'||p_value||'" '; END IF;
    if p_format_mask is not null then v_str := v_str||'format_mask="'||p_format_mask||'" '; end if;
    if p_header_align is not null then V_STR := V_STR||'header_align="'||lower(p_header_align)||'" '; end if;
    v_str := v_str||'>'; 
    
    return v_str;
  end bcoll;
  ------------------------------------------------------------------------------
  function ecoll(i BINARY_INTEGER) 
  return varchar2
  is
  begin
   return '</CELL>';
  end ecoll;
    ------------------------------------------------------------------------------
  function get_column_names(p_column_alias in apex_application_page_ir_col.column_alias%type)
  return APEX_APPLICATION_PAGE_IR_COL.report_label%TYPE
  is
  begin
    -- https://github.com/glebovpavel/IR_to_MSExcel/issues/9
    -- Thanks HeavyS 
    return  apex_plugin_util.replace_substitutions(p_value => l_report.column_names(p_column_alias));
  exception
    when others then
       raise_application_error(-20001,'get_column_names: p_column_alias='||p_column_alias||' '||SQLERRM);
  end get_column_names;
  ------------------------------------------------------------------------------
  function get_col_format_mask(p_column_alias in apex_application_page_ir_col.column_alias%type)
  return formatmask
  is
  begin
    return replace(l_report.col_format_mask(p_column_alias),'"','');
  exception
    when others then
       raise_application_error(-20001,'get_col_format_mask: p_column_alias='||p_column_alias||' '||SQLERRM);
  end get_col_format_mask;
  ------------------------------------------------------------------------------
  procedure set_col_format_mask(p_column_alias in apex_application_page_ir_col.column_alias%type,
                                p_format_mask  in formatmask)
  is
  begin
      l_report.col_format_mask(p_column_alias) := p_format_mask;
  exception
    when others then
       raise_application_error(-20001,'set_col_format_mask: p_column_alias='||p_column_alias||' '||SQLERRM);
  end set_col_format_mask;
  ------------------------------------------------------------------------------
  function get_header_alignment(p_column_alias in apex_application_page_ir_col.column_alias%type)
  return APEX_APPLICATION_PAGE_IR_COL.heading_alignment%TYPE
  is
  begin
    return l_report.header_alignment(p_column_alias);
  exception
    when others then
       raise_application_error(-20001,'get_header_alignment: p_column_alias='||p_column_alias||' '||SQLERRM);
  end get_header_alignment;
  ------------------------------------------------------------------------------
  function get_column_alignment(p_column_alias in apex_application_page_ir_col.column_alias%type)
  return apex_application_page_ir_col.column_alignment%type
  is
  begin
    return l_report.column_alignment(p_column_alias);
  exception
    when others then
       raise_application_error(-20001,'get_column_alignment: p_column_alias='||p_column_alias||' '||SQLERRM);
  end get_column_alignment;
  ------------------------------------------------------------------------------
  function get_column_types(p_num in binary_integer)
  return apex_application_page_ir_col.column_type%type
  is
  begin
    return l_report.column_types(p_num);
  exception
    when others then
       raise_application_error(-20001,'get_column_names: p_num='||p_num||' '||SQLERRM);
  END get_column_types;

  ------------------------------------------------------------------------------  
  function get_column_alias(p_num in binary_integer)
  return varchar2
  is
  begin
    return l_report.displayed_columns(p_num);
  exception
    when others then
       raise_application_error(-20001,'get_column_alias: p_num='||p_num||' '||SQLERRM);
  END get_column_alias;
  ------------------------------------------------------------------------------
  FUNCTION get_column_alias_sql(p_num IN binary_integer -- column number in sql-query
                               )
  return varchar2
  is
  BEGIN
    return l_report.displayed_columns(p_num - l_report.start_with + 1);
  exception
    WHEN others THEN
       raise_application_error(-20001,'get_column_alias_sql: p_num='||p_num||' '||SQLERRM);
  END get_column_alias_sql;
  ------------------------------------------------------------------------------
  function get_current_row(p_current_row in apex_application_global.vc_arr2,
                           p_id          in binary_integer)
  return largevarchar2
  is
  begin
    return p_current_row(p_id);
  exception
    when others then
       raise_application_error(-20001,'get_current_row: string: p_id='||p_id||' '||SQLERRM);
  end get_current_row; 
  ------------------------------------------------------------------------------
  function get_current_row(p_current_row in t_date_table,
                           p_id          in binary_integer)
  return date
  is
  begin
    if p_current_row.exists(p_id) then
      return p_current_row(p_id);
    else
      return null;
    end if;
  exception
    when others then
       raise_application_error(-20001,'get_current_row:date: p_id='||p_id||' '||SQLERRM);
  end get_current_row;   
  ------------------------------------------------------------------------------
  function get_current_row(p_current_row in t_number_table,
                           p_id          in binary_integer)
  return number
  is
  begin
    if p_current_row.exists(p_id) then
      return p_current_row(p_id);
    else
      return null;
    end if;
  exception
    when others then
       raise_application_error(-20001,'get_current_row:number: p_id='||p_id||' '||SQLERRM);
  end get_current_row;   

  ------------------------------------------------------------------------------
  -- :::: -> :
  function rr(p_str in varchar2)
  return varchar2
  is 
  begin
    return ltrim(rtrim(regexp_replace(p_str,'[:]+',':'),':'),':');
  end;
  ------------------------------------------------------------------------------   
 
  function get_xmlval(p_str in varchar2)
  return varchar2
  is
    v_tmp largevarchar2;
  begin
    -- p_str can be encoded html-string 
    -- wee need first convert to text
    v_tmp := REGEXP_REPLACE(p_str,'<(BR)\s*/*>',chr(13)||chr(10),1,0,'i');
    v_tmp := REGEXP_REPLACE(v_tmp,'<[^<>]+>',' ',1,0,'i');
    -- https://community.oracle.com/message/14074217#14074217
    v_tmp := regexp_replace(v_tmp, '[^[:print:]'||chr(13)||chr(10)||chr(9)||']', ' ');
    v_tmp := UTL_I18N.UNESCAPE_REFERENCE(v_tmp); 
    -- and finally encode them
    v_tmp := substr(v_tmp,1,2000);    
    v_tmp := UTL_I18N.ESCAPE_REFERENCE(v_tmp,'UTF8');
    return v_tmp;
  end get_xmlval;  
  ------------------------------------------------------------------------------
  
  function intersect_arrays(p_one IN APEX_APPLICATION_GLOBAL.VC_ARR2,
                            p_two IN APEX_APPLICATION_GLOBAL.VC_ARR2)
  return APEX_APPLICATION_GLOBAL.VC_ARR2
  is    
    v_ret APEX_APPLICATION_GLOBAL.VC_ARR2;
  begin    
    for i in 1..p_one.count loop
       for b in 1..p_two.count loop
         if p_one(i) = p_two(b) then
            v_ret(v_ret.count + 1) := p_one(i);
           exit;
         end if;
       end loop;        
    end loop;
    
    return v_ret;
  end intersect_arrays;
  ------------------------------------------------------------------------------
  function get_query_column_list
  return APEX_APPLICATION_GLOBAL.VC_ARR2
  is
   v_cur         INTEGER; 
   v_colls_count BINARY_INTEGER; 
   v_columns     APEX_APPLICATION_GLOBAL.VC_ARR2;
   v_desc_tab    DBMS_SQL.DESC_TAB2;
   v_sql         largevarchar2;   
  begin
    v_cur := dbms_sql.open_cursor(2);     
    v_sql := apex_plugin_util.replace_substitutions(p_value => l_report.report.sql_query,p_escape => false);
    log(v_sql);
    dbms_sql.parse(v_cur,v_sql,dbms_sql.native);     
    dbms_sql.describe_columns2(v_cur,v_colls_count,v_desc_tab);    
    for i in 1..v_colls_count loop
         if upper(v_desc_tab(i).col_name) != 'APXWS_ROW_PK' then --skip internal primary key if need
           v_columns(v_columns.count + 1) := v_desc_tab(i).col_name;
           log('Query column = '||v_desc_tab(i).col_name);
         end if;
    end loop;                 
   dbms_sql.close_cursor(v_cur);   
   return v_columns;
  exception
    when others then
      if dbms_sql.is_open(v_cur) then
        dbms_sql.close_cursor(v_cur);
      end if;  
      raise_application_error(-20001,'get_query_column_list: '||SQLERRM);
  end get_query_column_list;  
  ------------------------------------------------------------------------------
  function get_cols_as_table(p_delimetered_column_list     in varchar2,
                             p_displayed_nonbreak_columns  in apex_application_global.vc_arr2)
  return apex_application_global.vc_arr2
  is
  begin
    return intersect_arrays(APEX_UTIL.STRING_TO_TABLE(rr(p_delimetered_column_list)),p_displayed_nonbreak_columns);
  end get_cols_as_table;
  
  ------------------------------------------------------------------------------
  function get_hidden_columns_cnt(p_app_id       in number,
                                  p_page_id      in number,
                                  p_region_id    in number,
                                  p_report_id    in number)
  return number
  -- J.P.Lourens 9-Oct-16 added p_region_id as input variable, and added v_get_query_column_list
  is 
   v_hidden_columns  number default 0;
   v_hidden_computation_columns  number default 0;
   v_get_query_column_list varchar2(32676);
  begin
  
      v_get_query_column_list := apex_util.table_to_string(get_query_column_list);
      
      select count(*)
      into v_hidden_columns
      from APEX_APPLICATION_PAGE_IR_COL
     where application_id = p_app_id
       AND page_id = p_page_id
       -- J.P.Lourens 9-Oct-16 added p_region_id to ensure correct results when having multiple IR on a page
       and region_id = p_region_id
       and (display_text_as = 'HIDDEN'
       -- J.P.Lourens 9-Oct-2016 modified get_hidden_columns_cnt to INCLUDE columns which are
       --                        - selected in the IR query, and thus included in v_get_query_column_list
       --                        - not included in the report, thus missing from l_report.ir_data.report_columns
          or instr(':'||l_report.ir_data.report_columns||':',':'||column_alias||':') = 0)
       and instr(':'||v_get_query_column_list||':',':'||column_alias||':') > 0;
       
       select count(*)
       into v_hidden_computation_columns
       from apex_application_page_ir_comp
       where application_id = p_app_id
         and page_id = p_page_id       
         and report_id = p_report_id         
         and instr(':'||l_report.ir_data.report_columns||':',':'||computation_column_alias||':') = 0;       
      
      return v_hidden_columns + v_hidden_computation_columns;
  exception
    when no_data_found then
      return 0;
  end get_hidden_columns_cnt;     
  
  ------------------------------------------------------------------------------ 
  
  procedure init_t_report(p_app_id       IN NUMBER,
                          p_page_id      IN NUMBER,
                          p_region_id    IN NUMBER)
  is
    l_report_id     number;
    v_query_targets apex_application_global.vc_arr2;
    l_new_report    ir_report; 
  begin
    l_report := l_new_report;
    --get base report id    
    log('l_region_id='||p_region_id);
    
    l_report_id := apex_ir.get_last_viewed_report_id (p_page_id   => p_page_id,
                                                      p_region_id => p_region_id);
    
    log('l_base_report_id='||l_report_id);    
    
    select r.* 
    into l_report.ir_data       
    from apex_application_page_ir_rpt r
    where application_id = p_app_id 
      and page_id = p_page_id
      and session_id = v('APP_SESSION')
      and application_user = v('APP_USER')
      and base_report_id = l_report_id;
  
    log('l_report_id='||l_report_id);
    l_report_id := l_report.ir_data.report_id;                                                                 
      
      
    l_report.report := apex_ir.get_report (p_page_id        => p_page_id,
                                           p_region_id      => p_region_id
                                          );
    l_report.ir_data.report_columns := APEX_UTIL.TABLE_TO_STRING(get_cols_as_table(l_report.ir_data.report_columns,get_query_column_list));
    
    -- J.P.Lourens 9-Oct-16 added p_region_id as input variable
    l_report.hidden_cols_cnt := get_hidden_columns_cnt(p_app_id,p_page_id,p_region_id,l_report_id);
    
    <<displayed_columns>>                                      
    for i in (select column_alias,
                     report_label,
                     heading_alignment,
                     column_alignment,
                     column_type,
                     format_mask as  computation_format_mask,
                     nvl(instr(':'||l_report.ir_data.report_columns||':',':'||column_alias||':'),0) column_order ,
                     nvl(instr(':'||l_report.ir_data.break_enabled_on||':',':'||column_alias||':'),0) break_column_order
                from APEX_APPLICATION_PAGE_IR_COL
               where application_id = p_app_id
                 AND page_id = p_page_id
                 and region_id = p_region_id
                 and display_text_as != 'HIDDEN' --l_report.ir_data.report_columns can include HIDDEN columns
                 and instr(':'||l_report.ir_data.report_columns||':',':'||column_alias||':') > 0
              UNION
              select computation_column_alias,
                     computation_report_label,
                     'center' as heading_alignment,
                     'right' AS column_alignment,
                     computation_column_type,
                     computation_format_mask,
                     nvl(instr(':'||l_report.ir_data.report_columns||':',':'||computation_column_alias||':'),0) column_order,
                     nvl(instr(':'||l_report.ir_data.break_enabled_on||':',':'||computation_column_alias||':'),0) break_column_order
              from apex_application_page_ir_comp
              where application_id = p_app_id
                and page_id = p_page_id
                and report_id = l_report_id
                AND instr(':'||l_report.ir_data.report_columns||':',':'||computation_column_alias||':') > 0
              order by  break_column_order asc,column_order asc)
    loop                 
      l_report.column_names(i.column_alias) := i.report_label; 
      l_report.col_format_mask(i.column_alias) := i.computation_format_mask;
      l_report.header_alignment(i.column_alias) := i.heading_alignment; 
      l_report.column_alignment(i.column_alias) := i.column_alignment; 
      --l_report.column_types(i.column_alias) := i.column_type;
      IF i.column_order > 0 THEN
        IF i.break_column_order = 0 THEN 
          --displayed column
          l_report.displayed_columns(l_report.displayed_columns.count + 1) := i.column_alias;
        ELSE  
          --break column
          l_report.break_really_on(l_report.break_really_on.count + 1) := i.column_alias;
        end if;
      end if;  
      
      log('column='||i.column_alias||' l_report.column_names='||i.report_label);
      log('column='||i.column_alias||' l_report.col_format_mask='||i.computation_format_mask);
      log('column='||i.column_alias||' l_report.header_alignment='||i.heading_alignment);
      log('column='||i.column_alias||' l_report.column_alignment='||i.column_alignment);      
    end loop displayed_columns;    
    
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
    LOG('l_report.min_columns_on_break='||rr(l_report.ir_data.min_columns_on_break));
    log('l_report.median_columns_on_break='||rr(l_report.ir_data.median_columns_on_break));
    log('l_report.count_columns_on_break='||rr(l_report.ir_data.count_columns_on_break));    
    log('l_report.count_distnt_col_on_break='||rr(l_report.ir_data.count_distnt_col_on_break));
    log('l_report.break_really_on='||APEX_UTIL.TABLE_TO_STRING(l_report.break_really_on));
    log('l_report.agg_cols_cnt='||l_report.agg_cols_cnt);
    log('l_report.hidden_cols_cnt='||l_report.hidden_cols_cnt);
    
    
    for c in cur_highlight(p_report_id => l_report_id,
                           p_delimetered_column_list => l_report.ir_data.report_columns
                          ) 
    loop
        if c.HIGHLIGHT_ROW_COLOR is not null or c.HIGHLIGHT_ROW_FONT_COLOR is not null then
          --is row highlight
          l_report.row_highlight(l_report.row_highlight.count + 1) := c;        
        else
          l_report.col_highlight(l_report.col_highlight.count + 1) := c;           
        end if;  
        v_query_targets(v_query_targets.count + 1) := c.condition_sql||' as HLIGHTS_'||(v_query_targets.count + 1);
    end loop;    
        
    if v_query_targets.count  > 0 then
      -- uwr485kv is random name 
      l_report.report.sql_query := 'SELECT '||APEX_UTIL.TABLE_TO_STRING(v_query_targets,','||chr(10))||', uwr485kv.* from ('||l_report.report.sql_query||') uwr485kv';
    end if;
    l_report.report.sql_query := l_report.report.sql_query;
    log('l_report.report.sql_query='||chr(10)||l_report.report.sql_query||chr(10));
  exception
    when no_data_found then
      raise_application_error(-20001,'No Interactive Report found on Page='||p_page_id||' Application='||p_app_id||' Please make sure that the report was running at least once by this session.');
    when others then 
     log('Exception in init_t_report');
     log(' Page='||p_page_id);
     log(' Application='||p_app_id);
     raise;
  end init_t_report;  
  ------------------------------------------------------------------------------
 
  function is_control_break(p_curr_row  IN APEX_APPLICATION_GLOBAL.VC_ARR2,
                            p_prev_row  IN APEX_APPLICATION_GLOBAL.VC_ARR2)
  return boolean
  is
    v_start_with      BINARY_INTEGER;
    v_end_with        BINARY_INTEGER;    
  begin
    if nvl(l_report.break_really_on.count,0) = 0  then
      return false; --no control break
    end if;
    v_start_with := 1 + 
                    l_report.skipped_columns + 
                    l_report.row_highlight.count + 
                    l_report.col_highlight.count;    
    v_end_with   := l_report.skipped_columns + 
                    l_report.row_highlight.count + 
                    l_report.col_highlight.count +
                    nvl(l_report.break_really_on.count,0);
    for i in v_start_with..v_end_with loop
      if p_curr_row(i) != p_prev_row(i) then
        return true;
      end if;
    end loop;
    return false;
  end is_control_break;
  ------------------------------------------------------------------------------

  function get_formatted_number(p_number         IN number,
                                p_format_string  IN varchar2,
                                p_nls            IN varchar2 default null)
  return varchar2
  is
    v_str varchar2(100);
  begin
    v_str := trim(to_char(p_number,p_format_string,p_nls));
    if instr(v_str,'#') > 0 and ltrim(v_str,'#') is null then --format fail
      raise invalid_number;
    else
      return v_str;
    end if;    
  end get_formatted_number;  
  ------------------------------------------------------------------------------

  FUNCTION get_cell(p_query_value IN varchar2,
                    p_format_mask IN varchar2,
                    p_date        IN date)
  RETURN t_cell_data
  IS
    v_data t_cell_data;
  BEGIN     
     v_data.value := get_formatted_number(p_date - to_date('01-03-1900','DD-MM-YYYY') + 61,'9999999999999990D00000000','NLS_NUMERIC_CHARACTERS = ''.,''');     
     v_data.datatype := 'DATE';
     
     -- https://github.com/glebovpavel/IR_to_MSExcel/issues/16
     -- thanks Valentine Nikitsky 
     if p_format_mask is not null then
       if upper(p_format_mask)='SINCE' then
            v_data.text := apex_util.get_since(p_date);
            v_data.value := null;
            v_data.datatype := 'STRING';
       else
          v_data.text := to_char(p_date,p_format_mask);  -- normally v_data.text used in XML only
       end if;
     else
       v_data.text := p_query_value;
     end if;
     
     return v_data;
  EXCEPTION
    WHEN invalid_number or format_error THEN 
      v_data.value := NULL;          
      v_data.datatype := 'STRING';
      v_data.text := p_query_value;
      return v_data;
  END get_cell;  
  ------------------------------------------------------------------------------

  FUNCTION get_cell(p_query_value IN varchar2,
                    p_format_mask IN varchar2,
                    p_number      IN number)
  RETURN t_cell_data
  IS
    v_data t_cell_data;
  BEGIN
   v_data.datatype := 'NUMBER';   
   v_data.value := get_formatted_number(p_number,'9999999999999990D00000000','NLS_NUMERIC_CHARACTERS = ''.,''');
   
   if p_format_mask is not null then
     v_data.text := get_formatted_number(p_query_value,p_format_mask);
   ELSE
     v_data.text := p_query_value;
   end if;
   
   return v_data;
  EXCEPTION
    WHEN invalid_number or conversion_error THEN 
      v_data.value := NULL;          
      v_data.datatype := 'STRING';
      v_data.text := p_query_value;
      return v_data;
  END get_cell;  
  ------------------------------------------------------------------------------
  function print_row(p_current_row          IN apex_application_global.vc_arr2,
                     p_cur_date_row         IN t_date_table,
                     p_cur_number_row       IN t_number_table
                    )
  return varchar2 is
    v_clob            largevarchar2; --change
    v_column_alias    APEX_APPLICATION_PAGE_IR_COL.column_alias%TYPE;
    v_format_mask     APEX_APPLICATION_PAGE_IR_COMP.computation_format_mask%TYPE;
    v_row_color       varchar2(10); 
    v_row_back_color  varchar2(10);
    v_cell_color      varchar2(10);
    v_cell_back_color VARCHAR2(10);     
    v_column_type     columntype;
    v_cell_data       t_cell_data;
  begin
      --check that row need to be highlighted
    <<row_highlights>>
    for h in 1..l_report.row_highlight.count loop
     BEGIN 
      -- J.P.Lourens 9-Oct-16 
      -- current_row is based on report_sql which starts with the highlight columns, then the skipped columns and then the rest
      -- So to capture the highlight values the value for l_report.skipped_columns should NOT be taken into account
      IF get_current_row(p_current_row,/*l_report.skipped_columns + */l_report.row_highlight(h).COND_NUMBER) IS NOT NULL THEN
         v_row_color       := l_report.row_highlight(h).HIGHLIGHT_ROW_FONT_COLOR;
         v_row_back_color  := l_report.row_highlight(h).HIGHLIGHT_ROW_COLOR;
      END IF;
     exception       
       when no_data_found then
         log('row_highlights: ='||' end_with='||l_report.end_with||' agg_cols_cnt='||l_report.agg_cols_cnt||' COND_NUMBER='||l_report.row_highlight(h).cond_number||' h='||h);
     end; 
    end loop row_highlights;
    --
    <<visible_columns>>
    for i in l_report.start_with..l_report.end_with loop
      v_cell_color       := NULL;
      v_cell_back_color  := NULL;
      v_cell_data.value  := NULL;  
      v_cell_data.text   := NULL; 
      v_column_alias     := get_column_alias_sql(i);
      v_column_type      := get_column_types(i);
      v_format_mask      := get_col_format_mask(v_column_alias);
      
      IF v_column_type = 'DATE' THEN
         v_cell_data := get_cell(get_current_row(p_current_row,i),v_format_mask,get_current_row(p_cur_date_row,i));
      ELSIF  v_column_type = 'NUMBER' THEN      
         v_cell_data := get_cell(get_current_row(p_current_row,i),v_format_mask,get_current_row(p_cur_number_row,i));
      ELSE --STRING
        v_format_mask := NULL;
        v_cell_data.VALUE  := NULL;  
        v_cell_data.datatype := 'STRING';
        v_cell_data.text   := get_current_row(p_current_row,i);
      end if; 
       
      --check that cell need to be highlighted
      <<cell_highlights>>
      for h in 1..l_report.col_highlight.count loop
        begin
          -- J.P.Lourens 9-Oct-16 
          -- current_row is based on report_sql which starts with the highlight columns, then the skipped columns and then the rest
          -- So to capture the highlight values the value for l_report.skipped_columns should NOT be taken into account
          if get_current_row(p_current_row,/*l_report.skipped_columns + */l_report.col_highlight(h).COND_NUMBER) is not null 
             and v_column_alias = l_report.col_highlight(h).CONDITION_COLUMN_NAME 
          then
            v_cell_color       := l_report.col_highlight(h).HIGHLIGHT_CELL_FONT_COLOR;
            v_cell_back_color  := l_report.col_highlight(h).HIGHLIGHT_CELL_COLOR;
          end if;
        exception
       when no_data_found then
         log('col_highlights: ='||' end_with='||l_report.end_with||' agg_cols_cnt='||l_report.agg_cols_cnt||' COND_NUMBER='||l_report.col_highlight(h).cond_number||' h='||h); 
       end;
      END loop cell_highlights;
      
      v_clob := v_clob ||bcoll(p_font_color   => nvl(v_cell_color,v_row_color),
                               p_back_color   => nvl(v_cell_back_color,v_row_back_color),
                               p_align        => get_column_alignment(v_column_alias),
                               p_column_alias => v_column_alias,
                               p_colmn_type   => v_cell_data.datatype,
                               p_value        => v_cell_data.value,
                               p_format_mask  => v_format_mask
                              )
                       ||get_xmlval(v_cell_data.text)
                       ||ecoll(i);
    end loop visible_columns;
    return  '<ROW>'||v_clob || '</ROW>'||chr(10);    
  end print_row;
  
  ------------------------------------------------------------------------------  
  function print_header
  return varchar2 is
    v_header_xml      largevarchar2;
    v_column_alias    APEX_APPLICATION_PAGE_IR_COL.column_alias%TYPE;
  begin
    v_header_xml := '<HEADER>';
    <<headers>>
    for i in 1..l_report.displayed_columns.count   loop
      v_column_alias := get_column_alias(i);
      -- if current column is not control break column
      if apex_plugin_util.get_position_in_list(l_report.break_on,v_column_alias) is null then      
        v_header_xml := v_header_xml ||bcoll(p_column_alias => v_column_alias,
                                             p_header_align => get_header_alignment(v_column_alias),
                                             p_align        => get_column_alignment(v_column_alias),
                                             p_colmn_type   => get_column_types(i),
                                             p_format_mask  => get_col_format_mask(v_column_alias)
                                             )
                                     ||get_xmlval(regexp_replace(get_column_names(v_column_alias),'<[^>]*>',' '))
                                     ||ecoll(i);
      end if;  
    end loop headers;
    return  v_header_xml || '</HEADER>'||chr(10);
  end print_header; 
  ------------------------------------------------------------------------------  
  function print_control_break_header(p_current_row     in apex_application_global.vc_arr2) 
  return varchar2
  is
    v_cb_xml  largevarchar2;
  begin
    if nvl(l_report.break_really_on.count,0) = 0  then
      return ''; --no control break
    end if;
    
    <<break_columns>>
    for i in 1..nvl(l_report.break_really_on.count,0) loop
      --TODO: Add column header
      v_cb_xml := v_cb_xml ||
                  get_column_names(l_report.break_really_on(i))||': '||
                  get_current_row(p_current_row,i + 
                                                l_report.skipped_columns + 
                                                l_report.row_highlight.count + 
                                                l_report.col_highlight.count
                                 )||',';
    end loop break_columns;
    
    return  '<BREAK_HEADER>'||get_xmlval(rtrim(v_cb_xml,',')) || '</BREAK_HEADER>'||chr(10);    
  end print_control_break_header;
  ------------------------------------------------------------------------------
  function find_rel_position (p_curr_col_name    IN varchar2,
                              p_agg_rows         IN APEX_APPLICATION_GLOBAL.VC_ARR2)
  return BINARY_INTEGER
  is
    v_relative_position BINARY_INTEGER;
  begin
    <<aggregated_rows>>
    for i in 1..p_agg_rows.count loop
      if p_curr_col_name = p_agg_rows(i) then        
         return i;
      end if;
    end loop aggregated_rows;
    
    return null;
  end find_rel_position;
  ------------------------------------------------------------------------------
  function get_agg_text(p_curr_col_name   IN varchar2,
                        p_agg_rows        IN APEX_APPLICATION_GLOBAL.VC_ARR2,
                        p_current_row     IN APEX_APPLICATION_GLOBAL.VC_ARR2,
                        p_agg_text        IN varchar2,
                        p_position        in BINARY_INTEGER, --start position in sql-query
                        p_col_number      IN BINARY_INTEGER, --column position when displayed
                        p_default_format_mask     IN varchar2 default null,
                        p_overwrite_format_mask   IN varchar2 default null) --should be used forcibly  
  return varchar2
  is
    v_tmp_pos       BINARY_INTEGER;  -- current column position in sql-query 
    v_format_mask   apex_application_page_ir_comp.computation_format_mask%type;
    v_agg_value     largevarchar2;
    v_row_value     largevarchar2;
    v_g_format_mask varchar2(100);  
    v_col_alias     varchar2(255);
  begin
      v_tmp_pos := find_rel_position (p_curr_col_name,p_agg_rows); 
      if v_tmp_pos is not null then
        v_col_alias := get_column_alias_sql(p_col_number);
        v_g_format_mask :=  get_col_format_mask(v_col_alias);   
        v_format_mask := nvl(v_g_format_mask,p_default_format_mask);
        v_format_mask := nvl(p_overwrite_format_mask,v_format_mask);
        v_row_value :=  get_current_row(p_current_row,p_position + l_report.hidden_cols_cnt + v_tmp_pos);
        v_agg_value := trim(to_char(v_row_value,v_format_mask));

        log('--find an aggregate');
        log('p_col_number='||p_col_number);
        log('v_col_alias='||v_col_alias);
        log('v_g_format_mask='||v_g_format_mask);
        log('p_default_format_mask='||p_default_format_mask);
        log('v_tmp_pos='||v_tmp_pos);
        log('p_position='||p_position);        
        log('v_row_value='||v_row_value);      
        
        return  get_xmlval(p_agg_text||v_agg_value||' '||chr(10));
      else
        return  '';
      end if;    
  exception
     when others then
        log('!Exception in get_agg_text');
        log('p_col_number='||p_col_number);
        log('v_col_alias='||v_col_alias);
        log('v_g_format_mask='||v_g_format_mask);
        log('p_default_format_mask='||p_default_format_mask);
        log('v_tmp_pos='||v_tmp_pos);
        log('p_position='||p_position);
        -- J.P.Lourens 9-Oct-16 added to log             
        log('v_row_value='||v_row_value);
        log('v_format_mask='||v_format_mask);
        raise;
  end get_agg_text;
  ------------------------------------------------------------------------------
  function print_aggregate(p_current_row     IN APEX_APPLICATION_GLOBAL.VC_ARR2) 
  return varchar2
  is
    v_aggregate_xml   largevarchar2;
    v_position        BINARY_INTEGER;    
  begin
    if l_report.agg_cols_cnt  = 0 then
      return ''; --no aggregate
    end if;    
    v_aggregate_xml := '<AGGREGATE>';   
      
    
    <<visible_columns>>
    for i in l_report.start_with..l_report.end_with loop
      v_position := l_report.end_with; --aggregate are placed after displayed columns and computations
      v_aggregate_xml := v_aggregate_xml || bcoll(p_column_alias=>get_column_alias_sql(i),
                                                  --p_value => v_sum_value,
                                                  p_format_mask => get_col_format_mask(get_column_alias_sql(i))
                                                  );
      -- and second to XML-tag text to display as text concatenated with other aggregates
      -- one column cah have only one aggregate of each type
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.sum_columns_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => ' ',
                                       p_position      => v_position,
                                       p_col_number    => i);
      v_position := v_position + l_report.sum_columns_on_break.count;
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.avg_columns_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => 'Avgerage:',
                                       p_position      => v_position,
                                       p_col_number    => i,
                                       p_default_format_mask   => '999G999G999G999G990D000');
      v_position := v_position + l_report.avg_columns_on_break.count;                                       
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.max_columns_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => 'Max:',
                                       p_position      => v_position,
                                       p_col_number    => i);
      v_position := v_position + l_report.max_columns_on_break.count;                                 
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.min_columns_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => 'Min:',
                                       p_position      => v_position,
                                       p_col_number    => i);
      v_position := v_position + l_report.min_columns_on_break.count;                                 
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.median_columns_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => 'Median:',
                                       p_position      => v_position,
                                       p_col_number    => i,
                                       p_default_format_mask   => '999G999G999G999G990D000');
      v_position := v_position + l_report.median_columns_on_break.count;                                 
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.count_columns_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => 'Count:',
                                       p_position      => v_position,
                                       p_col_number    => i,
                                       p_overwrite_format_mask   => '999G999G999G999G990');
      v_position := v_position + l_report.count_columns_on_break.count;                                 
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.count_distnt_col_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => 'Count distinct:',
                                       p_position      => v_position,
                                       p_col_number    => i,
                                       p_overwrite_format_mask   => '999G999G999G999G990');
      v_aggregate_xml := v_aggregate_xml || ecoll(i);
    end loop visible_columns;
    return  v_aggregate_xml || '</AGGREGATE>'||chr(10);
  end print_aggregate;      
  ------------------------------------------------------------------------------
  function can_show_as_date(p_format_string in varchar2)
  return boolean
  is 
    v_dummy varchar2(50);
  begin
    v_dummy := to_char(sysdate,p_format_string);
    
    return true;
  exception
    when invalid_number or format_error or date_format_error or conversion_error then 
      return false;
  end can_show_as_date;  
  ------------------------------------------------------------------------------
  function get_current_format(p_data_type in binary_integer)
  return varchar2
  is
    v_format    formatmask;
    v_parameter varchar2(50);
  begin
    if p_data_type in (dbms_types.TYPECODE_TIMESTAMP_TZ,181) then
       v_parameter := 'NLS_TIMESTAMP_TZ_FORMAT';
    elsif p_data_type in (dbms_types.TYPECODE_TIMESTAMP_LTZ,231) then
       v_parameter := 'NLS_TIMESTAMP_TZ_FORMAT';
    elsif p_data_type in (dbms_types.TYPECODE_TIMESTAMP,180) then
       v_parameter := 'NLS_TIMESTAMP_FORMAT';      
    elsif p_data_type = dbms_types.TYPECODE_DATE then
       v_parameter := 'NLS_DATE_FORMAT';      
    else 
       return 'dd.mm.yyyy';
    end if;     
    
    SELECT value 
    into v_format
    FROM V$NLS_Parameters 
    WHERE parameter = v_parameter;    
    
    return v_format;
  end get_current_format;
  ------------------------------------------------------------------------------
  -- excel has only DATE-format mask (no timezones etc)
  -- if format mask can be shown in excel - show column as date- type
  -- else - as string 
  procedure prepare_col_format_mask(p_col_number in binary_integer,
                                    p_data_type  in binary_integer )
  is
    v_format_mask          formatmask;
    v_default_format_mask  formatmask;
    v_final_format_mask    formatmask;
    v_col_alias            varchar2(255);    
  begin
     v_default_format_mask := get_current_format(p_data_type);
     log('v_default_format_mask='||v_default_format_mask);
     begin
       v_col_alias   := get_column_alias_sql(p_col_number);
       v_format_mask := get_col_format_mask(v_col_alias);
       log('v_col_alias='||v_col_alias||' v_format_mask='||v_format_mask);
     exception
       when no_data_found then 
          v_format_mask := '';
     end;  
     
     v_final_format_mask := nvl(v_format_mask,v_default_format_mask);
     if can_show_as_date(v_final_format_mask) then
        log('Can show as date');
        if v_col_alias is not null then 
           set_col_format_mask(p_column_alias => v_col_alias,
                               p_format_mask  => v_final_format_mask);
        end if;                       
        l_report.column_types(p_col_number) := 'DATE';
     else
        log('Can not show as date');
        l_report.column_types(p_col_number) := 'STRING';
     end if;
  end prepare_col_format_mask;
  ------------------------------------------------------------------------------
  
  procedure get_xml_from_ir(v_data in out nocopy clob,p_max_rows in integer)
  is
   v_cur                INTEGER; 
   v_result             INTEGER;
   v_colls_count        BINARY_INTEGER;
   v_row                APEX_APPLICATION_GLOBAL.VC_ARR2;
   v_date_row           t_date_table;
   v_number_row         t_number_table;
   v_char_dummy         varchar2(1);
   v_date_dummy         date;
   v_number_dummy       number;
   v_prev_row           APEX_APPLICATION_GLOBAL.VC_ARR2;
   v_columns            APEX_APPLICATION_GLOBAL.VC_ARR2;
   v_current_row        number default 0;
   v_desc_tab           DBMS_SQL.DESC_TAB2;
   v_inside             boolean default false;
   v_sql                largevarchar2;
   v_bind_variables     DBMS_SQL.VARCHAR2_TABLE;
   v_buffer             largevarchar2;
   v_bind_var_name      varchar2(255);
   v_binded             boolean; 
   v_format_mask        formatmask;
  begin
    v_cur := dbms_sql.open_cursor(2); 
    v_sql := apex_plugin_util.replace_substitutions(p_value  => l_report.report.sql_query,
                                                    p_escape => false);    
    dbms_sql.parse(v_cur,v_sql,dbms_sql.native);     
    dbms_sql.describe_columns2(v_cur,v_colls_count,v_desc_tab);    
    --skip internal primary key if need
    for i in 1..v_desc_tab.count loop      
      if lower(v_desc_tab(i).col_name) = 'apxws_row_pk' then
        l_report.skipped_columns := 1;
      end if;
    end loop;
    
    l_report.start_with := 1 + 
                           l_report.skipped_columns +
                           nvl(l_report.break_really_on.count,0) + 
                           l_report.row_highlight.count + 
                           l_report.col_highlight.count;
    l_report.end_with   := l_report.skipped_columns + 
                           nvl(l_report.break_really_on.count,0) + 
                           l_report.displayed_columns.count  + 
                           l_report.row_highlight.count + 
                           l_report.col_highlight.count;    
                           
    log('l_report.start_with='||l_report.start_with);
    log('l_report.end_with='||l_report.end_with);
    log('l_report.skipped_columns='||l_report.skipped_columns);

    -- init column datatypes and format masks
    for i in 1..v_desc_tab.count loop      
      log('column_type='||v_desc_tab(i).col_type);
      -- don't know why column types in dbms_sql.describe_columns2 do not correspond to types in dbms_types      
      if v_desc_tab(i).col_type in (dbms_types.TYPECODE_TIMESTAMP_TZ,181,
                                    dbms_types.TYPECODE_TIMESTAMP_LTZ,231,
                                    dbms_types.TYPECODE_TIMESTAMP,180,
                                    dbms_types.TYPECODE_DATE) 
      then
         prepare_col_format_mask(p_col_number => i,p_data_type => v_desc_tab(i).col_type);      
      elsif v_desc_tab(i).col_type = dbms_types.TYPECODE_NUMBER then
         l_report.column_types(i) := 'NUMBER';
      else
         l_report.column_types(i) := 'STRING';
         log('column_type='||v_desc_tab(i).col_type||' STRING');
      end if;        
    end loop;
    
    v_bind_variables := wwv_flow_utilities.get_binds(v_sql);
       
    add(v_data,v_buffer,print_header); 
    log('<<bind variables>>');
    
    <<bind_variables>>
    for i in 1..v_bind_variables.count loop      
      v_bind_var_name := ltrim(v_bind_variables(i),':');
      if v_bind_var_name = 'APXWS_MAX_ROW_CNT' then      
         -- remove max_rows
         DBMS_SQL.BIND_VARIABLE (v_cur,'APXWS_MAX_ROW_CNT',p_max_rows);    
         log('Bind variable ('||i||')'||v_bind_var_name||'<'||p_max_rows||'>');
      else
        v_binded := false; 
        --first look report bind variables (filtering, search etc)
        <<bind_report_variables>>
        for a in 1..l_report.report.binds.count loop
          if v_bind_var_name = l_report.report.binds(a).name then
             DBMS_SQL.BIND_VARIABLE (v_cur,v_bind_var_name,l_report.report.binds(a).value);
             log('Bind variable as report variable ('||i||')'||v_bind_var_name||'<'||l_report.report.binds(a).value||'>');
             v_binded := true;
             exit;
          end if;
        end loop bind_report_variables;
        -- substantive strings in sql-queries can have bind variables too
        -- these variables are not in v_report.binds
        -- and need to be binded separately
        if not v_binded then
          DBMS_SQL.BIND_VARIABLE (v_cur,v_bind_var_name,v(v_bind_var_name));
          log('Bind variable ('||i||')'||v_bind_var_name||'<'||v(v_bind_var_name)||'>');          
        end if;        
       end if; 
    end loop;          
  
    log('<<define_columns>>');    
    for i in 1..v_colls_count loop
       log('define column '||i);
       log('column type '||v_desc_tab(i).col_type);      
       if l_report.column_types(i) = 'DATE' then   
         dbms_sql.define_column(v_cur, i, v_date_dummy);
       elsif l_report.column_types(i) = 'NUMBER' then   
         dbms_sql.define_column(v_cur, i, v_number_dummy);
       else --STRING
         dbms_sql.define_column(v_cur, i, v_char_dummy,32767);
       end if;         
    end loop define_columns;    
    
    v_result := dbms_sql.execute(v_cur);         
   
    log('<<main_cycle>>');
    <<main_cycle>>
    LOOP 
         IF DBMS_SQL.FETCH_ROWS(v_cur)>0 THEN          
         log('<<fetch>>');
           -- get column values of the row 
            v_current_row := v_current_row + 1;
            <<query_columns>>
            for i in 1..v_colls_count loop               
               log('column type '||v_desc_tab(i).col_type);
               v_row(i) := ' ';
               v_date_row(i) := NULL;
               v_number_row(i) := NULL;
               if l_report.column_types(i) = 'DATE' then
                 dbms_sql.column_value(v_cur, i,v_date_row(i));
                 v_row(i) := to_char(v_date_row(i));
               elsif l_report.column_types(i) = 'NUMBER' then
                dbms_sql.column_value(v_cur, i,v_number_row(i));
                v_row(i) := to_char(v_number_row(i));
               else 
                 dbms_sql.column_value(v_cur, i,v_row(i));                 
               end if;  
            end loop;     
            --check control break
            if v_current_row > 1 then
             if is_control_break(v_row,v_prev_row) then                                             
               add(v_data,v_buffer,'</ROWSET>'||chr(10));
               v_inside := false;
             end if;
            end if;
            if not v_inside then
              add(v_data,v_buffer,'<ROWSET>'||chr(10));
              add(v_data,v_buffer,print_control_break_header(v_row));
              --print aggregates
              add(v_data,v_buffer,print_aggregate(v_row));
              v_inside := true;
            END IF;            --            
            <<query_columns>>
            for i in 1..v_colls_count loop
              v_prev_row(i) := v_row(i);                           
            end loop;                 
            add(v_data,v_buffer,print_row(p_current_row    => v_row,
                                          p_cur_date_row   => v_date_row,
                                          p_cur_number_row => v_number_row
                                         ));
         ELSE
           EXIT; 
         END IF; 
    END LOOP main_cycle;        
    if v_inside then
       add(v_data,v_buffer,'</ROWSET>');
       v_inside := false;
    end if;
   add(v_data,v_buffer,' ',TRUE); 
   dbms_sql.close_cursor(v_cur);   
  end get_xml_from_ir;
  ------------------------------------------------------------------------------
  procedure get_final_xml(p_clob           IN OUT NOCOPY CLOB,
                          p_app_id         IN NUMBER,
                          p_region_id      IN NUMBER,
                          p_page_id        IN NUMBER,
                          p_items_list     IN VARCHAR2,
                          p_get_page_items IN CHAR,
                          p_max_rows       IN NUMBER)
  is
   v_rows    apex_application_global.vc_arr2;
   v_buffer  largevarchar2;
  begin
    add(p_clob,v_buffer,'<?xml version="1.0" encoding="UTF-8"?>'||chr(10)||'<DOCUMENT>'||chr(10));    
    add(p_clob,v_buffer,'<DATA>'||chr(10),TRUE);   
    get_xml_from_ir(p_clob,p_max_rows);
    add(p_clob,v_buffer,'</DATA>'||chr(10));
    add(p_clob,v_buffer,'</DOCUMENT>'||chr(10),TRUE);  
  end get_final_xml;
  ------------------------------------------------------------------------------
  procedure download_file(p_data        in clob,
                          p_mime_header in varchar2,
                          p_file_name   in varchar2)
  is
    v_blob        blob;
    v_desc_offset PLS_INTEGER := 1;
    v_src_offset  PLS_INTEGER := 1;
    v_lang        PLS_INTEGER := 0;
    v_warning     PLS_INTEGER := 0;   
  begin
        dbms_lob.createtemporary(v_blob,true);
        dbms_lob.converttoblob(v_blob, p_data, dbms_lob.getlength(p_data), v_desc_offset, v_src_offset, dbms_lob.default_csid, v_lang, v_warning);
        sys.htp.init;
        sys.owa_util.mime_header(p_mime_header, FALSE );
        sys.htp.p('Content-length: ' || sys.dbms_lob.getlength( v_blob));
        sys.htp.p('Content-Disposition: attachment; filename="'||p_file_name||'"' );
        sys.htp.p('Cache-Control: must-revalidate, max-age=0');
        sys.htp.p('Expires: Thu, 01 Jan 1970 01:00:00 CET');
        sys.htp.p('Set-Cookie: GPV_DOWNLOAD_STARTED=1;');
        sys.owa_util.http_header_close;
        sys.wpg_docload.download_file( v_blob );
        dbms_lob.freetemporary(v_blob);
  exception
     when others then 
        raise_application_error(-20001,'Download file '||SQLERRM);
  end download_file;
  ------------------------------------------------------------------------------
  procedure set_collection(p_collection_name in varchar2,p_data in clob)
  is
   v_tmp char;
  begin
    IF apex_collection.collection_exists (p_collection_name) = FALSE
    THEN
      apex_collection.create_collection (p_collection_name);
    END IF;
   begin
     select '1' --clob001
     into v_tmp
     from apex_collections 
     where collection_name = p_collection_name
        and seq_id = 1;
        
     apex_collection.update_member ( p_collection_name => p_collection_name
                                    , p_seq            => 1
                                    , p_clob001        => p_data);
   exception
     when no_data_found then
      apex_collection.add_member ( p_collection_name => p_collection_name
                                 , p_clob001         => p_data );
       
   end;
  end set_collection;
  ------------------------------------------------------------------------------
  procedure get_report_xml(p_app_id          IN NUMBER,
                           p_page_id         IN NUMBER,     
                           p_region_id       IN NUMBER,
                           p_return_type     IN CHAR DEFAULT 'X', -- "Q" for debug information, "X" for XML-Data
                           p_get_page_items  IN CHAR DEFAULT 'N', -- Y,N - include page items in XML
                           p_items_list      IN VARCHAR2,         -- "," delimetered list of items that for including in XML
                           p_collection_name IN VARCHAR2,         -- name of APEX COLLECTION to save XML, when null - download as file
                           p_max_rows        IN NUMBER            -- maximum rows for export                            
                          )
  is
    v_data      clob;    
  begin  
    dbms_lob.trim (v_debug,0);    
    dbms_lob.createtemporary(v_data,true);
    --APEX_DEBUG_MESSAGE.ENABLE_DEBUG_MESSAGES(p_level => 7);
    log('version=1.6');
    log('p_app_id='||p_app_id);
    log('p_page_id='||p_page_id);
    log('p_region_id='||p_region_id);
    log('p_return_type='||p_return_type);
    log('p_get_page_items='||p_get_page_items);
    log('p_items_list='||p_items_list);
    log('p_collection_name='||p_collection_name);
    log('p_max_rows='||p_max_rows);        
    if p_return_type = 'Q' then  -- debug information                    
        begin        
          init_t_report(p_app_id,p_page_id,p_region_id);              
          get_final_xml(v_data,p_app_id,p_page_id,p_region_id,p_items_list,p_get_page_items,p_max_rows);          
          if p_collection_name is not null then              
            set_collection(upper(p_collection_name),v_data);
          end if;
        exception
          when others then
            log('Error in IR_TO_XML.get_report_xml '||sqlerrm||chr(10)||chr(10)||dbms_utility.format_error_backtrace);            
        end;        
        log(' ',TRUE);
        download_file(v_debug,'text/txt','log.txt');        
    elsif p_return_type = 'X' then --XML-Data
        init_t_report(p_app_id,p_page_id,p_region_id);    
        get_final_xml(v_data,p_app_id,p_page_id,p_region_id,p_items_list,p_get_page_items,p_max_rows);
        if p_collection_name is not null then  
          set_collection(upper(p_collection_name),v_data);
        else
          download_file(v_data,'application/xml','report_data.xml');
        end if;
    else
      raise_application_error(-20001,'Unknown parameter p_download_type='||p_return_type);
      dbms_lob.freetemporary(v_data);
    end if;
    dbms_lob.freetemporary(v_data);
  exception
   when others then
     raise_application_error(-20001,'get_report_xml:'||SQLERRM);
     raise;
  end get_report_xml; 
  ------------------------------------------------------------------------------
  function get_report_xml(p_app_id          IN NUMBER,
                          p_page_id         IN NUMBER,  
                          p_region_id       IN NUMBER,
                          p_get_page_items  IN CHAR DEFAULT 'N', -- Y,N - include page items in XML
                          p_items_list      IN VARCHAR2,         -- "," delimetered list of items that for including in XML
                          p_max_rows        IN NUMBER            -- maximum rows for export                            
                         )
  return xmltype                           
  is
    v_data      clob;    
  begin
    dbms_lob.trim (v_debug,0);
    dbms_lob.createtemporary(v_data,true, DBMS_LOB.CALL);
    log('p_app_id='||p_app_id);
    log('p_page_id='||p_page_id);
    log('p_get_page_items='||p_get_page_items);
    log('p_items_list='||p_items_list);
    log('p_max_rows='||p_max_rows);
    
    init_t_report(p_app_id,p_page_id,p_region_id);
    get_final_xml(v_data,p_app_id,p_page_id,p_region_id,p_items_list,p_get_page_items,p_max_rows);    
    
    return xmltype(v_data);    
  end get_report_xml; 
  
  ------------------------------------------------------------------------------
  /* 
    function to handle cases of 'in' and 'not in' conditions for highlights
   	used in cursor cur_highlight
    
    Author: Srihari Ravva
  */ 
  function get_highlight_in_cond_sql(p_condition_expression  in APEX_APPLICATION_PAGE_IR_COND.CONDITION_EXPRESSION%TYPE,
                                     p_condition_sql         in APEX_APPLICATION_PAGE_IR_COND.CONDITION_SQL%TYPE,
                                     p_condition_column_name in APEX_APPLICATION_PAGE_IR_COND.CONDITION_COLUMN_NAME%TYPE)
  return varchar2 
  is
    v_condition_sql_tmp     varchar2(32767);
	  v_condition_sql			varchar2(32767);
	  v_arr_cond_expr APEX_APPLICATION_GLOBAL.VC_ARR2;
	  v_arr_cond_sql APEX_APPLICATION_GLOBAL.VC_ARR2;	
  begin
    v_condition_sql := REPLACE(REPLACE(p_condition_sql,'#APXWS_HL_ID#','1'),'#APXWS_CC_EXPR#','"'||p_condition_column_name||'"');
	  v_condition_sql_tmp := SUBSTR(v_condition_sql,INSTR(v_condition_sql,'#'),INSTR(v_condition_sql,'#',-1)-INSTR(v_condition_sql,'#')+1);
	
    v_arr_cond_expr := APEX_UTIL.STRING_TO_TABLE(p_condition_expression,',');
	v_arr_cond_sql := APEX_UTIL.STRING_TO_TABLE(v_condition_sql_tmp,',');
	
    for i in 1..v_arr_cond_expr.count
	loop
		-- consider everything as varchar2
		-- 'in' and 'not in' highlight conditions are not possible for DATE columns from IR
		v_condition_sql := REPLACE(v_condition_sql,v_arr_cond_sql(i),''''||TO_CHAR(v_arr_cond_expr(i))||'''');
	end loop;
    return v_condition_sql;
  end get_highlight_in_cond_sql;  
  
begin
  dbms_lob.createtemporary(v_debug,true, DBMS_LOB.CALL);  
END IR_TO_XML;
/
