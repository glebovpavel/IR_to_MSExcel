/**********************************************
**
** Author: Pavel Glebov
** Date: 08-2014
**
** This all in one install script contains headrs and bodies of 3 packages
**
** IR_TO_XML.sql 
** AS_ZIP.sql  
** XML_TO_XSLX.sql
**
**********************************************/
CREATE OR REPLACE package ir_to_xml as    
  --ver 1.3.
  -- download interactive report as PDF
  PROCEDURE get_report_xml(p_app_id          IN NUMBER,
                           p_page_id         in number,                                
                           p_return_type     IN CHAR DEFAULT 'X', -- "Q" for debug information "X" for XML-Data
                           p_get_page_items  IN CHAR DEFAULT 'N', -- Y,N - include page items in XML
                           p_items_list      in varchar2,         -- "," delimetered list of items that for including in XML
                           p_collection_name IN VARCHAR2,         -- name of APEX COLLECTION to save XML, when null - download as file
                           p_max_rows        IN NUMBER            -- maximum rows for export                            
                          );
  
  --return debug information
  function get_log return clob;
  
  -- get XML 
  function get_report_xml(p_app_id          IN NUMBER,
                          p_page_id         in number,                                
                          p_get_page_items  IN CHAR DEFAULT 'N', -- Y,N - include page items in XML
                          p_items_list      in varchar2,         -- "," delimetered list of items that for including in XML
                          p_max_rows        IN NUMBER            -- maximum rows for export                            
                         )
  return xmltype;     
                              
END IR_TO_XML;

/


CREATE OR REPLACE package body ir_to_xml as   
  
  subtype largevarchar2 is varchar2(32000); 
 
  cursor cur_highlight(p_report_id in APEX_APPLICATION_PAGE_IR_RPT.REPORT_ID%TYPE,
                       p_delimetered_column_list in varchar2) 
  IS
  select replace(replace(replace(replace(condition_sql,'#APXWS_EXPR#',''''||CONDITION_EXPRESSION||''''),'#APXWS_EXPR2#',''''||CONDITION_EXPRESSION2||''''),'#APXWS_HL_ID#','1'),'#APXWS_CC_EXPR#','"'||CONDITION_COLUMN_NAME||'"')  condition_sql,
       CONDITION_COLUMN_NAME,
       CONDITION_ENABLED,
       HIGHLIGHT_ROW_COLOR,
       HIGHLIGHT_ROW_FONT_COLOR,
       HIGHLIGHT_CELL_COLOR,
       HIGHLIGHT_CELL_FONT_COLOR,
       rownum COND_NUMBER,
       'HIGHLIGHT_'||rownum COND_NAME
  from APEX_APPLICATION_PAGE_IR_COND
  where condition_type = 'Highlight'
    and report_id = p_report_id
    and condition_enabled = 'Yes'
    and instr(':'||p_delimetered_column_list||':',':'||CONDITION_COLUMN_NAME||':') > 0
    order by --rows highlights first 
           nvl2(HIGHLIGHT_ROW_COLOR,1,0) desc, 
           nvl2(HIGHLIGHT_ROW_FONT_COLOR,1,0) desc,
           HIGHLIGHT_SEQUENCE;
  
  type t_col_names is table of apex_application_page_ir_col.report_label%type index by apex_application_page_ir_col.column_alias%type;
  type t_col_format_mask is table of APEX_APPLICATION_PAGE_IR_COMP.computation_format_mask%TYPE index by APEX_APPLICATION_PAGE_IR_COL.column_alias%TYPE;
  type t_header_alignment is table of APEX_APPLICATION_PAGE_IR_COL.heading_alignment%TYPE index by APEX_APPLICATION_PAGE_IR_COL.column_alias%TYPE;
  type t_column_alignment is table of apex_application_page_ir_col.column_alignment%type index by apex_application_page_ir_col.column_alias%type;
  type t_column_types is table of apex_application_page_ir_col.column_type%type index by apex_application_page_ir_col.column_alias%type;
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
    skipped_columns           INTEGER default 0, -- when scpecial coluns like apxws_row_pk is used
    start_with                INTEGER default 0, -- position of first displayed column in query
    end_with                  INTEGER default 0, -- position of last displayed column in query
    agg_cols_cnt              INTEGER default 0, 
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
     text            CLOB
   );  

  l_report    ir_report;   
  v_debug     clob;
  ------------------------------------------------------------------------------
  function get_log
  return clob
  is
  begin
    return v_debug;
  end  get_log;
  ------------------------------------------------------------------------------
  procedure add(p_clob in out nocopy clob,p_str varchar2)
  is
  begin
    if p_str is not null then
      dbms_lob.writeappend(p_clob,length(p_str),p_str);
    end if;  
  end;
  ------------------------------------------------------------------------------
  procedure log(p_message in varchar2)
  is
  begin
    add(v_debug,p_message||chr(10));
    apex_debug_message.log_message(p_message => substr(p_message,1,32767),
                                   p_enabled => false,
                                   p_level   => 4);
  end log; 
  ------------------------------------------------------------------------------
  function bcoll(p_font_color    in varchar2 default null,
                 p_back_color    in varchar2 default null,
                 p_align         in varchar2 default null,
                 p_width         in varchar2 default null,
                 p_column_alias  IN VARCHAR2 DEFAULT NULL,
                 p_colmn_type    IN VARCHAR2 DEFAULT NULL,
                 p_value         IN VARCHAR2 DEFAULT NULL,
                 p_format_mask   IN VARCHAR2 DEFAULT NULL) 
  return varchar2
  is
    v_str varchar2(500);
  begin
    v_str := v_str||'<CELL ';
    if p_column_alias is not null then v_str := v_str||'column-alias="'||p_column_alias||'" '; end if;
    if p_font_color is not null then v_str := v_str||'color="'||p_font_color||'" '; end if;
    if p_colmn_type is not null then V_STR := V_STR||'data-type="'||p_colmn_type||'" '; end if;
    if p_back_color is not null then v_str := v_str||'background-color="'||p_back_color||'" '; end if;
    if p_align is not null then V_STR := V_STR||'align="'||lower(p_align)||'" '; end if;
    IF p_width IS NOT NULL THEN v_str := v_str||'width="'||p_width||'" '; END IF;        
    IF p_value IS NOT NULL THEN v_str := v_str||'value="'||p_value||'" '; END IF;
    if p_format_mask is not null then v_str := v_str||'format_mask="'||p_format_mask||'" '; end if;
    v_str := v_str||'>'; 
    
    return v_str;
  end bcoll;
  ------------------------------------------------------------------------------
  function ecoll(i integer) 
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
    return l_report.column_names(p_column_alias);
  exception
    when others then
       raise_application_error(-20001,'get_column_names: p_column_alias='||p_column_alias||' '||SQLERRM);
  end get_column_names;
  ------------------------------------------------------------------------------
  function get_col_format_mask(p_column_alias in apex_application_page_ir_col.column_alias%type)
  return APEX_APPLICATION_PAGE_IR_COMP.computation_format_mask%TYPE
  is
  begin
    return l_report.col_format_mask(p_column_alias);
  exception
    when others then
       raise_application_error(-20001,'get_column_names: p_column_alias='||p_column_alias||' '||SQLERRM);
  end get_col_format_mask;
  ------------------------------------------------------------------------------
  function get_header_alignment(p_column_alias in apex_application_page_ir_col.column_alias%type)
  return APEX_APPLICATION_PAGE_IR_COL.heading_alignment%TYPE
  is
  begin
    return l_report.header_alignment(p_column_alias);
  exception
    when others then
       raise_application_error(-20001,'get_column_names: p_column_alias='||p_column_alias||' '||SQLERRM);
  end get_header_alignment;
  ------------------------------------------------------------------------------
  function get_column_alignment(p_column_alias in apex_application_page_ir_col.column_alias%type)
  return apex_application_page_ir_col.column_alignment%type
  is
  begin
    return l_report.column_alignment(p_column_alias);
  exception
    when others then
       raise_application_error(-20001,'get_column_names: p_column_alias='||p_column_alias||' '||SQLERRM);
  end get_column_alignment;
  ------------------------------------------------------------------------------
  function get_column_types(p_column_alias in apex_application_page_ir_col.column_alias%type)
  return apex_application_page_ir_col.column_type%type
  is
  begin
    return l_report.column_types(p_column_alias);
  exception
    when others then
       raise_application_error(-20001,'get_column_names: p_column_alias='||p_column_alias||' '||SQLERRM);
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
                           p_id in binary_integer)
  return apex_application_page_ir_col.column_type%type
  is
  begin
    return p_current_row(p_id);
  exception
    when others then
       raise_application_error(-20001,'get_current_row: p_id='||p_id||' '||SQLERRM);
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
  begin
    return dbms_xmlgen.convert(p_str);
    --RETURN REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(p_str,'<','%26lt;'),'>','%26gt;'),'&','%26amp;'),'"','%26quot;'),'''','%26apos;');
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
  function intersect_count(p_one IN APEX_APPLICATION_GLOBAL.VC_ARR2,
                           p_two IN APEX_APPLICATION_GLOBAL.VC_ARR2)
  return integer
  is
   v_rez APEX_APPLICATION_GLOBAL.VC_ARR2;
  begin
    v_rez := intersect_arrays(p_one,p_two);
    return v_rez.count;
  end intersect_count; 
  ------------------------------------------------------------------------------
  
  procedure init_t_report(p_app_id       in number,
                          p_page_id      in number)
  is
    l_region_id     number;
    l_report_id     number;
    v_query_targets apex_application_global.vc_arr2;
    l_new_report    ir_report;
  begin
    l_report := l_new_report;

    select region_id 
    into l_region_id 
    from APEX_APPLICATION_PAGE_REGIONS 
    where application_id = p_app_id 
      and page_id = p_page_id 
      and source_type = 'Interactive Report';    
    
    --get base report id    
    log('l_region_id='||l_region_id);
    
    l_report_id := apex_ir.get_last_viewed_report_id (p_page_id   => p_page_id,
                                                      p_region_id => l_region_id);
    
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
                                           p_region_id      => l_region_id
                                           --p_report_id      => l_report_id
                                          );
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
                 and display_text_as != 'HIDDEN' --after report RESETTING l_report.ir_data.report_columns consists HIDDEN column - APEX bug????
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
      l_report.column_types(i.column_alias) := i.column_type;
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
      log('column='||i.column_alias||' l_report.column_types='||i.column_type);
    end loop;    

    --l_report.break_on := APEX_UTIL.STRING_TO_TABLE(rr(l_report.ir_data.break_enabled_on));    
    l_report.sum_columns_on_break := APEX_UTIL.STRING_TO_TABLE(rr(l_report.ir_data.sum_columns_on_break));  
    l_report.avg_columns_on_break := APEX_UTIL.STRING_TO_TABLE(rr(l_report.ir_data.avg_columns_on_break));  
    l_report.max_columns_on_break := APEX_UTIL.STRING_TO_TABLE(rr(l_report.ir_data.max_columns_on_break));  
    l_report.min_columns_on_break := APEX_UTIL.STRING_TO_TABLE(rr(l_report.ir_data.min_columns_on_break));  
    l_report.median_columns_on_break := APEX_UTIL.STRING_TO_TABLE(rr(l_report.ir_data.median_columns_on_break)); 
    l_report.count_columns_on_break := APEX_UTIL.STRING_TO_TABLE(rr(l_report.ir_data.count_columns_on_break));  
    l_report.count_distnt_col_on_break := APEX_UTIL.STRING_TO_TABLE(rr(l_report.ir_data.count_distnt_col_on_break)); 
      
    l_report.agg_cols_cnt := l_report.sum_columns_on_break.count + 
                             l_report.avg_columns_on_break.count +
                             l_report.max_columns_on_break.count + 
                             l_report.min_columns_on_break.count +
                             l_report.median_columns_on_break.count +
                             l_report.count_columns_on_break.count +
                             l_report.count_distnt_col_on_break.count;
    
    log('l_report.displayed_columns='||rr(l_report.ir_data.report_columns));
    log('l_report.break_on='||rr(l_report.ir_data.break_enabled_on));
    log('l_report.sum_columns_on_break='||rr(l_report.ir_data.sum_columns_on_break));
    log('l_report.avg_columns_on_break='||rr(l_report.ir_data.avg_columns_on_break));
    log('l_report.max_columns_on_break='||rr(l_report.ir_data.max_columns_on_break));
    LOG('l_report.min_columns_on_break='||rr(l_report.ir_data.min_columns_on_break));
    log('l_report.median_columns_on_break='||rr(l_report.ir_data.median_columns_on_break));
    log('l_report.count_columns_on_break='||rr(l_report.ir_data.count_distnt_col_on_break));
    log('l_report.count_distnt_col_on_break='||rr(l_report.ir_data.count_columns_on_break));
    log('l_report.break_really_on='||APEX_UTIL.TABLE_TO_STRING(l_report.break_really_on));
    log('l_report.agg_cols_cnt='||l_report.agg_cols_cnt);
    
    v_query_targets(v_query_targets.count + 1) := 'rez.*';
     
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
        v_query_targets(v_query_targets.count + 1) := c.condition_sql;
    end loop;    
        
    l_report.report.sql_query := 'SELECT '||APEX_UTIL.TABLE_TO_STRING(v_query_targets,',')||' from ( '
                                          ||l_report.report.sql_query||' ) rez';
    log('l_report.report.sql_query='||chr(10)||l_report.report.sql_query||chr(10));
  exception
    when no_data_found then
      raise_application_error(-20001,'No Interactive Report found on Page='||p_page_id||' Application='||p_app_id||' Please make sure that the report was running at least once by this session.');
    when others then 
      raise_application_error(-20001,'init_t_report: Page='||p_page_id||' Application='||p_app_id||' '||sqlerrm);
  end init_t_report;  
  ------------------------------------------------------------------------------
 
  function is_control_break(p_curr_row  IN APEX_APPLICATION_GLOBAL.VC_ARR2,
                            p_prev_row  IN APEX_APPLICATION_GLOBAL.VC_ARR2)
  return boolean
  is
    v_start_with      integer;
    v_end_with        integer;    
    v_tmp             integer;
  begin
    if nvl(l_report.break_really_on.count,0) = 0  then
      return false; --no control break
    end if;
    v_start_with := 1 + l_report.skipped_columns;    
    v_end_with   := l_report.skipped_columns + nvl(l_report.break_really_on.count,0);
    for i in v_start_with..v_end_with loop
      if p_curr_row(i) != p_prev_row(i) then
        return true;
      end if;
    end loop;
    return false;
  end is_control_break;
  ------------------------------------------------------------------------------
  FUNCTION get_cell_date(p_query_value IN VARCHAR2,p_format_mask IN VARCHAR2)
  RETURN t_cell_data
  IS
    v_data t_cell_data;
  BEGIN
     BEGIN
       v_data.value := to_date(p_query_value) - to_date('01-03-1900','DD-MM-YYYY') + 61;
       if p_format_mask is not null then
         v_data.text := to_char(to_date(p_query_value),p_format_mask);
       ELSE
         v_data.text := p_query_value;
       end if;
      exception
        WHEN others THEN 
          v_data.text := p_query_value;
      END;      
      
      return v_data;
  end get_cell_date;
  ------------------------------------------------------------------------------
  FUNCTION get_cell_number(p_query_value IN VARCHAR2,p_format_mask IN VARCHAR2)
  RETURN t_cell_data
  IS
    v_data t_cell_data;
  BEGIN
     begin
       v_data.value := trim(to_char(to_number(p_query_value),'9999999999999999999999990D0000000000000000000000000','NLS_NUMERIC_CHARACTERS = ''.,'''));
       
       if p_format_mask is not null then
         v_data.text := trim(to_char(to_number(p_query_value),p_format_mask));
       ELSE
         v_data.text := p_query_value;
       end if;
      exception
        WHEN others THEN 
          v_data.text := p_query_value;
      END;      
      
      return v_data;
  END get_cell_number;  
  ------------------------------------------------------------------------------
  function print_row(p_current_row     IN APEX_APPLICATION_GLOBAL.VC_ARR2)
  return varchar2 is
    v_clob            largevarchar2; --change
    v_column_alias    APEX_APPLICATION_PAGE_IR_COL.column_alias%TYPE;
    v_format_mask     APEX_APPLICATION_PAGE_IR_COMP.computation_format_mask%TYPE;
    v_row_color       varchar2(10); 
    v_row_back_color  varchar2(10);
    v_cell_color      varchar2(10);
    v_cell_back_color VARCHAR2(10);     
    v_column_type     VARCHAR2(10);
    v_cell_data       t_cell_data;
  begin
      --check that row need to be highlighted
    <<row_highlights>>
    for h in 1..l_report.row_highlight.count loop
     BEGIN 
      IF get_current_row(p_current_row,l_report.end_with + l_report.agg_cols_cnt + l_report.row_highlight(h).COND_NUMBER) IS NOT NULL THEN
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
      v_cell_data.VALUE  := NULL;  
      v_cell_data.text   := NULL; 

      v_column_alias := get_column_alias_sql(i);
      v_column_type := get_column_types(v_column_alias);
      v_format_mask := get_col_format_mask(v_column_alias);
      
      IF v_column_type = 'DATE' THEN
         v_cell_data := get_cell_date(get_current_row(p_current_row,i),v_format_mask);                   
      ELSIF  v_column_type = 'NUMBER' THEN      
         v_cell_data := get_cell_number(get_current_row(p_current_row,i),v_format_mask);
      ELSE --STRING
        v_format_mask := NULL;
        v_cell_data.VALUE  := NULL;  
        v_cell_data.text   := get_current_row(p_current_row,i);
      end if; 
       
      --check that cell need to be highlighted
      <<column_highlights>>
      for h in 1..l_report.col_highlight.count loop
        begin
          --debug
          if get_current_row(p_current_row,l_report.end_with + l_report.agg_cols_cnt + l_report.col_highlight(h).COND_NUMBER) is not null 
             and v_column_alias = l_report.col_highlight(h).CONDITION_COLUMN_NAME 
          then
            v_cell_color       := l_report.col_highlight(h).HIGHLIGHT_CELL_FONT_COLOR;
            v_cell_back_color  := l_report.col_highlight(h).HIGHLIGHT_CELL_COLOR;
          end if;
        exception
       when no_data_found then
         log('col_highlights: ='||' end_with='||l_report.end_with||' agg_cols_cnt='||l_report.agg_cols_cnt||' COND_NUMBER='||l_report.col_highlight(h).cond_number||' h='||h); 
       end;
      END loop column_highlights;
      
      v_clob := v_clob ||bcoll(p_font_color   => nvl(v_cell_color,v_row_color),
                               p_back_color   => nvl(v_cell_back_color,v_row_back_color),
                               p_align        => get_column_alignment(v_column_alias),
                               p_column_alias => v_column_alias,
                               p_colmn_type   => v_column_type,
                               p_value        => v_cell_data.value,
                               p_format_mask  => replace(v_format_mask,'"','')
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
      V_COLUMN_ALIAS := get_column_alias(i);
      -- if current column is not control break column
      if apex_plugin_util.get_position_in_list(l_report.break_on,v_column_alias) is null then      
        v_header_xml := v_header_xml ||bcoll(p_column_alias=>v_column_alias,p_align=>get_header_alignment(v_column_alias))
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
      v_cb_xml := v_cb_xml ||get_column_names(l_report.break_really_on(i))||': '||get_current_row(p_current_row,i + l_report.skipped_columns)||',';
    end loop visible_columns;
    
    return  '<BREAK_HEADER>'||get_xmlval(rtrim(v_cb_xml,',')) || '</BREAK_HEADER>'||chr(10);    
  end print_control_break_header;
  ------------------------------------------------------------------------------
  function find_rel_position (p_curr_col_name    IN varchar2,
                              p_agg_rows         IN APEX_APPLICATION_GLOBAL.VC_ARR2)
  return integer
  is
    v_relative_position integer;
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
                        p_position        in integer, --start position in sql-query
                        p_col_number      IN INTEGER, --column position when displayed
                        p_default_format_mask     IN varchar2 default null )  
  return varchar2
  is
    v_tmp_pos       integer;  -- current column position in sql-query 
    v_format_mask   apex_application_page_ir_comp.computation_format_mask%type;
    v_agg_value     varchar2(1000);
  begin
      v_tmp_pos := find_rel_position (p_curr_col_name,p_agg_rows); 
      if v_tmp_pos is not null then
        v_format_mask := nvl(get_col_format_mask(get_column_alias_sql(p_col_number)),p_default_format_mask);
        v_agg_value := trim(to_char(get_current_row(p_current_row,p_position + v_tmp_pos),v_format_mask));
        
        return  get_xmlval(p_agg_text||v_agg_value||' '||chr(10));
      else
        return  '';
      end if;        
  end get_agg_text;
  ------------------------------------------------------------------------------
  function get_agg_value(p_curr_col_name   in varchar2,
                         p_agg_rows        IN APEX_APPLICATION_GLOBAL.VC_ARR2,
                         p_current_row     in apex_application_global.vc_arr2,
                         p_position        in integer --start position in sql-query
                        )  
  return varchar2
  is
    v_tmp_pos       integer;  -- current column position in sql-query 
    v_format_mask   apex_application_page_ir_comp.computation_format_mask%type;
    v_agg_value     varchar2(100);
  begin
      v_tmp_pos := find_rel_position (p_curr_col_name,p_agg_rows); 
      if v_tmp_pos is not null then
        v_agg_value := get_current_row(p_current_row,p_position + v_tmp_pos);
        return  get_xmlval(v_agg_value);
      else
        return  '';
      end if;        
  end get_agg_value;
  
  ------------------------------------------------------------------------------
  function print_aggregate(p_current_row     IN APEX_APPLICATION_GLOBAL.VC_ARR2) 
  return varchar2
  is
    v_aggregate_xml   largevarchar2;
    v_position        integer;    
    v_sum_value       varchar2(100);
  begin
    if l_report.agg_cols_cnt  = 0 then
      return ''; --no aggregate
    end if;    
    v_aggregate_xml := '<AGGREGATE>';   
      
    
    <<visible_columns>>
    for i in l_report.start_with..l_report.end_with loop
      v_position := l_report.end_with; --aggregate are placed after displayed columns and computations
      v_sum_value := get_agg_value(p_curr_col_name => get_column_alias_sql(i),
                         p_agg_rows      => l_report.sum_columns_on_break,
                         p_current_row   => p_current_row,
                         p_position      => v_position
                        );
                        
      v_aggregate_xml := v_aggregate_xml || bcoll(p_column_alias=>get_column_alias_sql(i),
                                                  p_value => v_sum_value,
                                                  p_format_mask => get_col_format_mask(get_column_alias_sql(i))
                                                  );
      --one column cah have only one aggregate of each type
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
                                       p_col_number    => i);
      v_position := v_position + l_report.count_columns_on_break.count;                                 
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.count_distnt_col_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => 'Count distinct:',
                                       p_position      => v_position,
                                       p_col_number    => i);
      v_aggregate_xml := v_aggregate_xml || ecoll(i);
    end loop visible_columns;
    return  v_aggregate_xml || '</AGGREGATE>'||chr(10);
  end print_aggregate;    
  ------------------------------------------------------------------------------
  function get_page_items(p_app_id         in number,
                          p_page_id        in number,
                          p_items_list     in varchar2,
                          p_get_page_items in char)
  return clob
  is
    v_clob  clob;    
    v_item_names  APEX_APPLICATION_GLOBAL.VC_ARR2;
  begin
    v_clob := to_clob( '<ITEMS>'||chr(10));
    
    select item_name
    bulk collect into v_item_names  
    from apex_application_page_items
    where application_id = p_app_id
      and ((page_id = p_page_id and p_get_page_items = 'Y')
          or
          (P_ITEMS_LIST is not null and INSTR(','||P_ITEMS_LIST||',',','||ITEM_NAME||',') >  0))
    union 
    select item_name
    from APEX_APPLICATION_ITEMS
    where application_id = p_app_id  
      and P_ITEMS_LIST is not null 
      and instr(','||p_items_list||',',','||item_name||',') >  0;    
    
    <<items>>
    for i in 1..v_item_names.count loop
     v_clob := v_clob||to_clob('<'||upper(v_item_names(i))||'>'
                                ||get_xmlval(v(v_item_names(i)))
                                ||'</'||upper(v_item_names(i))||'>'||chr(10));
    end loop items;
    
    return v_clob||to_clob('</ITEMS>'||chr(10)); 
  end get_page_items;  
 
  ------------------------------------------------------------------------------    
  procedure get_xml_from_ir(v_data in out nocopy clob,p_max_rows in integer)
  is
   v_cur         INTEGER; 
   v_result      INTEGER;
   v_colls_count INTEGER;
   v_row         APEX_APPLICATION_GLOBAL.VC_ARR2;
   v_prev_row    APEX_APPLICATION_GLOBAL.VC_ARR2;
   v_columns     APEX_APPLICATION_GLOBAL.VC_ARR2;
   v_current_row number default 0;
   v_desc_tab    DBMS_SQL.DESC_TAB2;
   v_inside      boolean default false;
  begin
    v_cur := dbms_sql.open_cursor; 
    dbms_sql.parse(v_cur,l_report.report.sql_query,dbms_sql.native);     
    dbms_sql.describe_columns2(v_cur,v_colls_count,v_desc_tab);    
    --skip internal primary key if need
    if lower(v_desc_tab(1).col_name) = 'apxws_row_pk' then
      l_report.skipped_columns := 1;
    end if;
    l_report.start_with := l_report.skipped_columns + 1 + nvl(l_report.break_really_on.count,0);
    l_report.end_with   := l_report.skipped_columns + nvl(l_report.break_really_on.count,0) + l_report.displayed_columns.count;    
    log('l_report.start_with='||l_report.start_with);
    log('l_report.end_with='||l_report.end_with);
    log('l_report.skipped_columns='||l_report.skipped_columns);
    
    add(v_data,print_header); 
    
    <<bind_variables>>
    for i in 1..l_report.report.binds.count loop
      --remove MAX_ROWS
      if l_report.report.binds(i).name = 'APXWS_MAX_ROW_CNT' then      
        DBMS_SQL.BIND_VARIABLE (v_cur,l_report.report.binds(i).name,p_max_rows);      
        null;
      else
        DBMS_SQL.BIND_VARIABLE (v_cur,l_report.report.binds(i).name,l_report.report.binds(i).value);      
      end if;
    end loop bind_variables;

    <<query_columns>>
    for i in 1..v_colls_count loop
     v_row(i) := '';
     dbms_sql.define_column(v_cur, i, v_row(i),32767);
    end loop query_columns;
    
    v_result := dbms_sql.execute(v_cur);         
    <<main_cycle>>
    LOOP 
         IF DBMS_SQL.FETCH_ROWS(v_cur)>0 THEN          
           -- get column values of the row 
            v_current_row := v_current_row + 1;
            <<query_columns>>
            for i in 1..v_colls_count loop
               DBMS_SQL.COLUMN_VALUE(v_cur, i,v_row(i));                
            end loop;     
            --check control break
            if v_current_row > 1 then
             if is_control_break(v_row,v_prev_row) then                                             
               add(v_data,'</ROWSET>'||chr(10));
               v_inside := false;
             end if;
            end if;
            if not v_inside then
              add(v_data,'<ROWSET>'||chr(10));
              add(v_data,print_control_break_header(v_row));
              --print aggregates
              add(v_data,print_aggregate(v_row));
              v_inside := true;
            END IF;            --            
            <<query_columns>>
            for i in 1..v_colls_count loop
              v_prev_row(i) := v_row(i);                           
            end loop;                 
            --v_xml := v_xml||to_clob(print_row(v_row));
            add(v_data,print_row(v_row));
         ELSE --DBMS_SQL.FETCH_ROWS(v_cur)>0
           EXIT; 
         END IF; 
    END LOOP main_cycle;        
    if v_inside then
       add(v_data,'</ROWSET>');
       v_inside := false;
    end if;
   dbms_sql.close_cursor(v_cur); 
   
  end get_xml_from_ir;
  ------------------------------------------------------------------------------
  procedure get_final_xml( p_clob           in out nocopy clob,
                          p_app_id         in number,
                          p_page_id        in number,
                          p_items_list     in varchar2,
                          p_get_page_items in char,
                          p_max_rows       in number)
  is
   v_rows  apex_application_global.vc_arr2;
  begin
    add(p_clob,'<?xml version="1.0" encoding="UTF-8"?>'||chr(10)||'<DOCUMENT>'||chr(10));    
    add(p_clob,get_page_items(p_app_id,p_page_id,p_items_list,p_get_page_items));
    add(p_clob,'<DATA>'||chr(10));
   
    get_xml_from_ir(p_clob,p_max_rows);    
   
    add(p_clob,'</DATA>'||chr(10));
    add(p_clob,'</DOCUMENT>'||chr(10));
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
        sys.owa_util.http_header_close;
        sys.wpg_docload.download_file( v_blob );
        dbms_lob.freetemporary(v_blob);
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
                           p_page_id         in number,                                
                           p_return_type     IN CHAR DEFAULT 'X', -- "Q" for debug information, "X" for XML-Data
                           p_get_page_items  IN CHAR DEFAULT 'N', -- Y,N - include page items in XML
                           p_items_list      in varchar2,         -- "," delimetered list of items that for including in XML
                           p_collection_name IN VARCHAR2,         -- name of APEX COLLECTION to save XML, when null - download as file
                           p_max_rows        IN NUMBER            -- maximum rows for export                            
                          )
  is
    v_data      clob;
  begin
    dbms_lob.trim (v_debug,0);
    dbms_lob.createtemporary(v_data,true);
    --APEX_DEBUG_MESSAGE.ENABLE_DEBUG_MESSAGES(p_level => 7);
    log('p_app_id='||p_app_id);
    log('p_page_id='||p_page_id);
    log('p_return_type='||p_return_type);
    log('p_get_page_items='||p_get_page_items);
    log('p_items_list='||p_items_list);
    log('p_collection_name='||p_collection_name);
    log('p_max_rows='||p_max_rows);
    
    init_t_report(p_app_id,p_page_id);    
    if p_return_type = 'Q' then  -- debug information    
        begin
          get_final_xml(v_data,p_app_id,p_page_id,p_items_list,p_get_page_items,p_max_rows);
          if p_collection_name is not null then  
            set_collection(upper(p_collection_name),v_data);
          end if;
        exception
          when others then
            log('Error in IR_TO_XML.get_report_document '||sqlerrm||chr(10)||chr(10)||dbms_utility.format_error_backtrace);            
        end;
        download_file(v_debug,'text/txt','log.txt');
    elsif p_return_type = 'X' then --XML-Data
        get_final_xml(v_data,p_app_id,p_page_id,p_items_list,p_get_page_items,p_max_rows);
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
  end get_report_xml; 
  ------------------------------------------------------------------------------
  function get_report_xml(p_app_id          IN NUMBER,
                          p_page_id         in number,                                
                          p_get_page_items  IN CHAR DEFAULT 'N', -- Y,N - include page items in XML
                          p_items_list      in varchar2,         -- "," delimetered list of items that for including in XML
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
    
    init_t_report(p_app_id,p_page_id);
    get_final_xml(v_data,p_app_id,p_page_id,p_items_list,p_get_page_items,p_max_rows);    
    
    return xmltype(v_data);    
  end get_report_xml; 

begin
  dbms_lob.createtemporary(v_debug,true, DBMS_LOB.CALL);  
END IR_TO_XML;
/
CREATE OR REPLACE package as_zip
is
/**********************************************
**
** Author: Anton Scheffer
** Date: 25-01-2012
** Website: http://technology.amis.nl/blog
**
** Changelog:
**   Date: 29-04-2012
**    fixed bug for large uncompressed files, thanks Morten Braten
**   Date: 21-03-2012
**     Take CRC32, compressed length and uncompressed length from 
**     Central file header instead of Local file header
**   Date: 17-02-2012
**     Added more support for non-ascii filenames
**   Date: 25-01-2012
**     Added MIT-license
**     Some minor improvements
**
******************************************************************************
******************************************************************************
Copyright (C) 2010,2011 by Anton Scheffer

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

******************************************************************************
******************************************** */
  type file_list is table of clob;
--
  function file2blob
    ( p_dir varchar2
    , p_file_name varchar2
    )
  return blob;
--
  function get_file_list
    ( p_dir varchar2
    , p_zip_file varchar2
    , p_encoding varchar2 := null
    )
  return file_list;
--
  function get_file_list
    ( p_zipped_blob blob
    , p_encoding varchar2 := null
    )
  return file_list;
--
  function get_file
    ( p_dir varchar2
    , p_zip_file varchar2
    , p_file_name varchar2
    , p_encoding varchar2 := null
    )
  return blob;
--
  function get_file
    ( p_zipped_blob blob
    , p_file_name varchar2
    , p_encoding varchar2 := null
    )
  return blob;
--
  procedure add1file
    ( p_zipped_blob in out blob
    , p_name varchar2
    , p_content blob
    );
--
  procedure finish_zip( p_zipped_blob in out blob );
--
  procedure save_zip
    ( p_zipped_blob blob
    , p_dir varchar2 := 'MY_DIR'
    , p_filename varchar2 := 'my.zip'
    );
--
/*
declare
  g_zipped_blob blob;
begin
  as_zip.add1file( g_zipped_blob, 'test4.txt', null ); -- a empty file
  as_zip.add1file( g_zipped_blob, 'dir1/test1.txt', utl_raw.cast_to_raw( q'<A file with some more text, stored in a subfolder which isn't added>' ) );
  as_zip.add1file( g_zipped_blob, 'test1234.txt', utl_raw.cast_to_raw( 'A small file' ) );
  as_zip.add1file( g_zipped_blob, 'dir2/', null ); -- a folder
  as_zip.add1file( g_zipped_blob, 'dir3/', null ); -- a folder
  as_zip.add1file( g_zipped_blob, 'dir3/test2.txt', utl_raw.cast_to_raw( 'A small filein a previous created folder' ) );
  as_zip.finish_zip( g_zipped_blob );
  as_zip.save_zip( g_zipped_blob, 'MY_DIR', 'my.zip' );
  dbms_lob.freetemporary( g_zipped_blob );
end;
--
declare
  zip_files as_zip.file_list;
begin
  zip_files  := as_zip.get_file_list( 'MY_DIR', 'my.zip' );
  for i in zip_files.first() .. zip_files.last
  loop
    dbms_output.put_line( zip_files( i ) );
    dbms_output.put_line( utl_raw.cast_to_varchar2( as_zip.get_file( 'MY_DIR', 'my.zip', zip_files( i ) ) ) );
  end loop;
end;
*/
end;
/


CREATE OR REPLACE package body as_zip
is
--
  c_LOCAL_FILE_HEADER        constant raw(4) := hextoraw( '504B0304' ); -- Local file header signature
  c_END_OF_CENTRAL_DIRECTORY constant raw(4) := hextoraw( '504B0506' ); -- End of central directory signature
--
  function blob2num( p_blob blob, p_len integer, p_pos integer )
  return number
  is
  begin
    return utl_raw.cast_to_binary_integer( dbms_lob.substr( p_blob, p_len, p_pos ), utl_raw.little_endian );
  end;
--
  function raw2varchar2( p_raw raw, p_encoding varchar2 )
  return varchar2
  is
  begin
    return coalesce( utl_i18n.raw_to_char( p_raw, p_encoding )
                   , utl_i18n.raw_to_char( p_raw, utl_i18n.map_charset( p_encoding, utl_i18n.GENERIC_CONTEXT, utl_i18n.IANA_TO_ORACLE ) )
                   );
  end;
--
  function little_endian( p_big number, p_bytes pls_integer := 4 )
  return raw
  is
  begin
    return utl_raw.substr( utl_raw.cast_from_binary_integer( p_big, utl_raw.little_endian ), 1, p_bytes );
  end;
--
  function file2blob
    ( p_dir varchar2
    , p_file_name varchar2
    )
  return blob
  is
    file_lob bfile;
    file_blob blob;
  begin
    file_lob := bfilename( p_dir, p_file_name );
    dbms_lob.open( file_lob, dbms_lob.file_readonly );
    dbms_lob.createtemporary( file_blob, true );
    dbms_lob.loadfromfile( file_blob, file_lob, dbms_lob.lobmaxsize );
    dbms_lob.close( file_lob );
    return file_blob;
  exception
    when others then
      if dbms_lob.isopen( file_lob ) = 1
      then
        dbms_lob.close( file_lob );
      end if;
      if dbms_lob.istemporary( file_blob ) = 1
      then
        dbms_lob.freetemporary( file_blob );
      end if;
      raise;
  end;
--
  function get_file_list
    ( p_zipped_blob blob
    , p_encoding varchar2 := null
    )
  return file_list
  is
    t_ind integer;
    t_hd_ind integer;
    t_rv file_list;
    t_encoding varchar2(32767);
  begin
    t_ind := dbms_lob.getlength( p_zipped_blob ) - 21;
    loop
      exit when t_ind < 1 or dbms_lob.substr( p_zipped_blob, 4, t_ind ) = c_END_OF_CENTRAL_DIRECTORY;
      t_ind := t_ind - 1;
    end loop;
--
    if t_ind <= 0
    then
      return null;
    end if;
--
    t_hd_ind := blob2num( p_zipped_blob, 4, t_ind + 16 ) + 1;
    t_rv := file_list();
    t_rv.extend( blob2num( p_zipped_blob, 2, t_ind + 10 ) );
    for i in 1 .. blob2num( p_zipped_blob, 2, t_ind + 8 )
    loop
      if p_encoding is null
      then
        if utl_raw.bit_and( dbms_lob.substr( p_zipped_blob, 1, t_hd_ind + 9 ), hextoraw( '08' ) ) = hextoraw( '08' )
        then  
          t_encoding := 'AL32UTF8'; -- utf8
        else
          t_encoding := 'US8PC437'; -- IBM codepage 437
        end if;
      else
        t_encoding := p_encoding;
      end if;
      t_rv( i ) := raw2varchar2
                     ( dbms_lob.substr( p_zipped_blob
                                      , blob2num( p_zipped_blob, 2, t_hd_ind + 28 )
                                      , t_hd_ind + 46
                                      )
                     , t_encoding
                     );
      t_hd_ind := t_hd_ind + 46
                + blob2num( p_zipped_blob, 2, t_hd_ind + 28 )  -- File name length
                + blob2num( p_zipped_blob, 2, t_hd_ind + 30 )  -- Extra field length
                + blob2num( p_zipped_blob, 2, t_hd_ind + 32 ); -- File comment length
    end loop;
--
    return t_rv;
  end;
--
  function get_file_list
    ( p_dir varchar2
    , p_zip_file varchar2
    , p_encoding varchar2 := null
    )
  return file_list
  is
  begin
    return get_file_list( file2blob( p_dir, p_zip_file ), p_encoding );
  end;
--
  function get_file
    ( p_zipped_blob blob
    , p_file_name varchar2
    , p_encoding varchar2 := null
    )
  return blob
  is
    t_tmp blob;
    t_ind integer;
    t_hd_ind integer;
    t_fl_ind integer;
    t_encoding varchar2(32767);
    t_len integer;
  begin
    t_ind := dbms_lob.getlength( p_zipped_blob ) - 21;
    loop
      exit when t_ind < 1 or dbms_lob.substr( p_zipped_blob, 4, t_ind ) = c_END_OF_CENTRAL_DIRECTORY;
      t_ind := t_ind - 1;
    end loop;
--
    if t_ind <= 0
    then
      return null;
    end if;
--
    t_hd_ind := blob2num( p_zipped_blob, 4, t_ind + 16 ) + 1;
    for i in 1 .. blob2num( p_zipped_blob, 2, t_ind + 8 )
    loop
      if p_encoding is null
      then
        if utl_raw.bit_and( dbms_lob.substr( p_zipped_blob, 1, t_hd_ind + 9 ), hextoraw( '08' ) ) = hextoraw( '08' )
        then  
          t_encoding := 'AL32UTF8'; -- utf8
        else
          t_encoding := 'US8PC437'; -- IBM codepage 437
        end if;
      else
        t_encoding := p_encoding;
      end if;
      if p_file_name = raw2varchar2
                         ( dbms_lob.substr( p_zipped_blob
                                          , blob2num( p_zipped_blob, 2, t_hd_ind + 28 )
                                          , t_hd_ind + 46
                                          )
                         , t_encoding
                         )
      then
        t_len := blob2num( p_zipped_blob, 4, t_hd_ind + 24 ); -- uncompressed length 
        if t_len = 0
        then
          if substr( p_file_name, -1 ) in ( '/', '\' )
          then  -- directory/folder
            return null;
          else -- empty file
            return empty_blob();
          end if;
        end if;
--
        if dbms_lob.substr( p_zipped_blob, 2, t_hd_ind + 10 ) = hextoraw( '0800' ) -- deflate
        then
          t_fl_ind := blob2num( p_zipped_blob, 4, t_hd_ind + 42 );
          t_tmp := hextoraw( '1F8B0800000000000003' ); -- gzip header
          dbms_lob.copy( t_tmp
                       , p_zipped_blob
                       ,  blob2num( p_zipped_blob, 4, t_hd_ind + 20 )
                       , 11
                       , t_fl_ind + 31
                       + blob2num( p_zipped_blob, 2, t_fl_ind + 27 ) -- File name length
                       + blob2num( p_zipped_blob, 2, t_fl_ind + 29 ) -- Extra field length
                       );
          dbms_lob.append( t_tmp, utl_raw.concat( dbms_lob.substr( p_zipped_blob, 4, t_hd_ind + 16 ) -- CRC32
                                                , little_endian( t_len ) -- uncompressed length
                                                )
                         );
          return utl_compress.lz_uncompress( t_tmp );
        end if;
--
        if dbms_lob.substr( p_zipped_blob, 2, t_hd_ind + 10 ) = hextoraw( '0000' ) -- The file is stored (no compression)
        then
          t_fl_ind := blob2num( p_zipped_blob, 4, t_hd_ind + 42 );
          dbms_lob.createtemporary( t_tmp, true );
          dbms_lob.copy( t_tmp
                       , p_zipped_blob
                       , t_len
                       , 1
                       , t_fl_ind + 31
                       + blob2num( p_zipped_blob, 2, t_fl_ind + 27 ) -- File name length
                       + blob2num( p_zipped_blob, 2, t_fl_ind + 29 ) -- Extra field length
                       );
          return t_tmp;
        end if;
      end if;
      t_hd_ind := t_hd_ind + 46
                + blob2num( p_zipped_blob, 2, t_hd_ind + 28 )  -- File name length
                + blob2num( p_zipped_blob, 2, t_hd_ind + 30 )  -- Extra field length
                + blob2num( p_zipped_blob, 2, t_hd_ind + 32 ); -- File comment length
    end loop;
--
    return null;
  end;
--
  function get_file
    ( p_dir varchar2
    , p_zip_file varchar2
    , p_file_name varchar2
    , p_encoding varchar2 := null
    )
  return blob
  is
  begin
    return get_file( file2blob( p_dir, p_zip_file ), p_file_name, p_encoding );
  end;
--
  procedure add1file
    ( p_zipped_blob in out blob
    , p_name varchar2
    , p_content blob
    )
  is
    t_now date;
    t_blob blob;
    t_len integer;
    t_clen integer;
    t_crc32 raw(4) := hextoraw( '00000000' );
    t_compressed boolean := false;
    t_name raw(32767);
  begin
    t_now := sysdate;
    t_len := nvl( dbms_lob.getlength( p_content ), 0 );
    if t_len > 0
    then 
      t_blob := utl_compress.lz_compress( p_content );
      t_clen := dbms_lob.getlength( t_blob ) - 18;
      t_compressed := t_clen < t_len;
      t_crc32 := dbms_lob.substr( t_blob, 4, t_clen + 11 );       
    end if;
    if not t_compressed
    then 
      t_clen := t_len;
      t_blob := p_content;
    end if;
    if p_zipped_blob is null
    then
      dbms_lob.createtemporary( p_zipped_blob, true );
    end if;
    t_name := utl_i18n.string_to_raw( p_name, 'AL32UTF8' );
    dbms_lob.append( p_zipped_blob
                   , utl_raw.concat( c_LOCAL_FILE_HEADER -- Local file header signature
                                   , hextoraw( '1400' )  -- version 2.0
                                   , case when t_name = utl_i18n.string_to_raw( p_name, 'US8PC437' )
                                       then hextoraw( '0000' ) -- no General purpose bits
                                       else hextoraw( '0008' ) -- set Language encoding flag (EFS)
                                     end 
                                   , case when t_compressed
                                        then hextoraw( '0800' ) -- deflate
                                        else hextoraw( '0000' ) -- stored
                                     end
                                   , little_endian( to_number( to_char( t_now, 'ss' ) ) / 2
                                                  + to_number( to_char( t_now, 'mi' ) ) * 32
                                                  + to_number( to_char( t_now, 'hh24' ) ) * 2048
                                                  , 2
                                                  ) -- File last modification time
                                   , little_endian( to_number( to_char( t_now, 'dd' ) )
                                                  + to_number( to_char( t_now, 'mm' ) ) * 32
                                                  + ( to_number( to_char( t_now, 'yyyy' ) ) - 1980 ) * 512
                                                  , 2
                                                  ) -- File last modification date
                                   , t_crc32 -- CRC-32
                                   , little_endian( t_clen )                      -- compressed size
                                   , little_endian( t_len )                       -- uncompressed size
                                   , little_endian( utl_raw.length( t_name ), 2 ) -- File name length
                                   , hextoraw( '0000' )                           -- Extra field length
                                   , t_name                                       -- File name
                                   )
                   );
    if t_compressed
    then                   
      dbms_lob.copy( p_zipped_blob, t_blob, t_clen, dbms_lob.getlength( p_zipped_blob ) + 1, 11 ); -- compressed content
    elsif t_clen > 0
    then                   
      dbms_lob.copy( p_zipped_blob, t_blob, t_clen, dbms_lob.getlength( p_zipped_blob ) + 1, 1 ); --  content
    end if;
    if dbms_lob.istemporary( t_blob ) = 1
    then      
      dbms_lob.freetemporary( t_blob );
    end if;
  end;
--
  procedure finish_zip( p_zipped_blob in out blob )
  is
    t_cnt pls_integer := 0;
    t_offs integer;
    t_offs_dir_header integer;
    t_offs_end_header integer;
    t_comment raw(32767) := utl_raw.cast_to_raw( 'Implementation by Anton Scheffer' );
  begin
    t_offs_dir_header := dbms_lob.getlength( p_zipped_blob );
    t_offs := 1;
    while dbms_lob.substr( p_zipped_blob, utl_raw.length( c_LOCAL_FILE_HEADER ), t_offs ) = c_LOCAL_FILE_HEADER
    loop
      t_cnt := t_cnt + 1;
      dbms_lob.append( p_zipped_blob
                     , utl_raw.concat( hextoraw( '504B0102' )      -- Central directory file header signature
                                     , hextoraw( '1400' )          -- version 2.0
                                     , dbms_lob.substr( p_zipped_blob, 26, t_offs + 4 )
                                     , hextoraw( '0000' )          -- File comment length
                                     , hextoraw( '0000' )          -- Disk number where file starts
                                     , hextoraw( '0000' )          -- Internal file attributes => 
                                                                   --     0000 binary file
                                                                   --     0100 (ascii)text file
                                     , case
                                         when dbms_lob.substr( p_zipped_blob
                                                             , 1
                                                             , t_offs + 30 + blob2num( p_zipped_blob, 2, t_offs + 26 ) - 1
                                                             ) in ( hextoraw( '2F' ) -- /
                                                                  , hextoraw( '5C' ) -- \
                                                                  )
                                         then hextoraw( '10000000' ) -- a directory/folder
                                         else hextoraw( '2000B681' ) -- a file
                                       end                         -- External file attributes
                                     , little_endian( t_offs - 1 ) -- Relative offset of local file header
                                     , dbms_lob.substr( p_zipped_blob
                                                      , blob2num( p_zipped_blob, 2, t_offs + 26 )
                                                      , t_offs + 30
                                                      )            -- File name
                                     )
                     );
      t_offs := t_offs + 30 + blob2num( p_zipped_blob, 4, t_offs + 18 )  -- compressed size
                            + blob2num( p_zipped_blob, 2, t_offs + 26 )  -- File name length 
                            + blob2num( p_zipped_blob, 2, t_offs + 28 ); -- Extra field length
    end loop;
    t_offs_end_header := dbms_lob.getlength( p_zipped_blob );
    dbms_lob.append( p_zipped_blob
                   , utl_raw.concat( c_END_OF_CENTRAL_DIRECTORY                                -- End of central directory signature
                                   , hextoraw( '0000' )                                        -- Number of this disk
                                   , hextoraw( '0000' )                                        -- Disk where central directory starts
                                   , little_endian( t_cnt, 2 )                                 -- Number of central directory records on this disk
                                   , little_endian( t_cnt, 2 )                                 -- Total number of central directory records
                                   , little_endian( t_offs_end_header - t_offs_dir_header )    -- Size of central directory
                                   , little_endian( t_offs_dir_header )                        -- Offset of start of central directory, relative to start of archive
                                   , little_endian( nvl( utl_raw.length( t_comment ), 0 ), 2 ) -- ZIP file comment length
                                   , t_comment
                                   )
                   );
  end;
--
  procedure save_zip
    ( p_zipped_blob blob
    , p_dir varchar2 := 'MY_DIR'
    , p_filename varchar2 := 'my.zip'
    )
  is
    t_fh utl_file.file_type;
    t_len pls_integer := 32767;
  begin
    t_fh := utl_file.fopen( p_dir, p_filename, 'wb' );
    for i in 0 .. trunc( ( dbms_lob.getlength( p_zipped_blob ) - 1 ) / t_len )
    loop
      utl_file.put_raw( t_fh, dbms_lob.substr( p_zipped_blob, t_len, i * t_len + 1 ) );
    end loop;
    utl_file.fclose( t_fh );
  end;
--
end;
/
CREATE OR REPLACE package xml_to_xslx
-- ver 1.0.
IS
  procedure download_file(p_app_id in number,
                    p_page_id      in number,
                    p_max_rows     in number,
                    p_file_name    in varchar2 default 'Excel'); 
end;
/


CREATE OR REPLACE PACKAGE body XML_TO_XSLX
is
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
  
  cursor cur_row(p_xml xmltype) is 
  SELECT rownum coll_num,
                                 extractvalue(COLUMN_VALUE, 'CELL/@background-color') AS background_color,
                                 extractvalue(COLUMN_VALUE, 'CELL/@color') AS font_color, 
                                 extractvalue(COLUMN_VALUE, 'CELL/@data-type') AS data_type,
                                 extractvalue(COLUMN_VALUE, 'CELL/@value') AS cell_value,
                                 extractvalue(column_value, 'CELL') as cell_text
                          from table (select xmlsequence(extract(p_xml,'DOCUMENT/DATA/ROWSET/ROW/CELL')) from dual);
  
  ------------------------------------------------------------------------------
  procedure add(p_clob in out nocopy clob,p_str varchar2)
  is
  begin
    DBMS_LOB.WRITEAPPEND(p_clob,length(p_str),p_str);
  end;
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
  procedure get_excel(p_xml in xmltype,v_clob in out nocopy clob,v_strings_clob in out nocopy clob)
  IS
    v_strings          apex_application_global.vc_arr2;
    v_rownum           binary_integer default 1;
    v_colls_count      binary_integer default 0;
    v_agg_clob         clob;
    v_agg_strings_cnt  binary_integer default 1; 
    string_height      constant number default 14.4; 
    aggregate_style_id constant number default 5;
    HEADER_STYLE_ID    constant number default 6;
    --
    procedure print_char_cell(p_coll in binary_integer, p_row in binary_integer, p_string in varchar2,p_clob in out nocopy CLOB,p_style_id in number default null)
    is
     v_style varchar2(20);
    begin
      if p_style_id is not null then
       v_style := ' s="'||p_style_id||'" ';
      end if;
      
      add(p_clob,'<c r="'||get_cell_name(p_coll,p_row)||'" t="s" '||v_style||'>'||chr(10)
                         ||'<v>' || to_char(v_strings.count)|| '</v>'||chr(10)                 
                         ||'</c>'||chr(10));
      v_strings(v_strings.count + 1) := p_string;
    end print_char_cell;
    --
    procedure print_number_cell(p_coll in binary_integer, p_row in binary_integer, p_value in varchar2,p_clob in out nocopy clob,p_style_id in number default null)
    is
      v_style varchar2(20);
    begin
      if p_style_id is not null then
       v_style := ' s="'||p_style_id||'" ';
      end if;

      add(p_clob,'<c r="'||get_cell_name(p_coll,p_row)||'" '||v_style||'>'||chr(10)
                         ||'<v>'||p_value|| '</v>'||chr(10)
                         ||'</c>'||chr(10));
    
    end print_number_cell;
    --
  begin
     pragma inline(add,'YES');     
     pragma inline(get_cell_name,'YES');
     
     dbms_lob.createtemporary(v_agg_clob,true);
     
     add(v_clob,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'||chr(10));
     add(v_clob,'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <dimension ref="A1"/>
        <sheetViews>
          <sheetView tabSelected="1" workbookViewId="0">
            <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
            <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
           </sheetView>
          </sheetViews>
        <sheetFormatPr baseColWidth="10" defaultColWidth="10" defaultRowHeight="15"/>'||chr(10));
     add(v_clob,'<sheetData>'||chr(10));
     
     --header
     add(v_clob,'<row>'||chr(10));     
     for i in (select  extractvalue(column_value, 'CELL') as column_header
               from table (select xmlsequence(extract(p_xml,'DOCUMENT/DATA/HEADER/CELL')) from dual))
     loop          
       v_colls_count := v_colls_count + 1;
       print_char_cell(p_coll => v_colls_count, p_row => v_rownum, p_string => i.column_header,p_clob => v_clob);
     end loop; 
     v_rownum := v_rownum + 1;
     add(v_clob,'</row>'||chr(10));     
     
     <<rowset>>     
     for rowset_xml in (select column_value as rowset,
                               extractvalue(COLUMN_VALUE, 'ROWSET/BREAK_HEADER') AS break_header
                        from table (select xmlsequence(extract(p_xml,'DOCUMENT/DATA/ROWSET')) from dual)
                       ) 
     loop
       --header
       if rowset_xml.break_header is not null then
         add(v_clob,'<row>'||chr(10));
         print_char_cell(p_coll => 1,p_row => v_rownum,p_string => rowset_xml.break_header,p_clob => v_clob,p_style_id => header_style_id);         
         --for u in 2..v_colls_count loop
         --  print_char_cell(p_coll => 1,p_row => v_rownum,p_string => '',p_clob => v_clob,p_style_id => HEADER_STYLE_ID);
         --end loop;
         v_rownum := v_rownum + 1;
         add(v_clob,'</row>'||chr(10));
       end if;
       
       <<cells>>
       for row_xml in (select column_value as row_ from table (select xmlsequence(extract(rowset_xml.rowset,'ROWSET/ROW')) from dual)) loop
         add(v_clob,'<row>'||chr(10));
         FOR cell_xml IN (SELECT rownum coll_num,
                                 extractvalue(COLUMN_VALUE, 'CELL/@background-color') AS background_color,
                                 extractvalue(COLUMN_VALUE, 'CELL/@color') AS font_color, 
                                 extractvalue(COLUMN_VALUE, 'CELL/@data-type') AS data_type,
                                 extractvalue(COLUMN_VALUE, 'CELL/@value') AS cell_value,
                                 extractvalue(COLUMN_VALUE, 'CELL') AS cell_text
                          from table (select xmlsequence(extract(row_xml.row_,'ROW/CELL')) from dual))
          loop
            begin            
              if cell_xml.data_type in ('NUMBER') then
                  print_number_cell(p_coll => cell_xml.coll_num, p_row => v_rownum, p_value => cell_xml.cell_value,p_clob => v_clob);
              elsif cell_xml.data_type in ('DATE') then
                 add(v_clob,'<c r="'||get_cell_name(cell_xml.coll_num,v_rownum)||'"  s="4">'||chr(10)
                                    ||'<v>'||cell_xml.cell_value|| '</v>'||chr(10)
                                    ||'</c>'||chr(10));
              else --STRING
                  print_char_cell(p_coll => cell_xml.coll_num,p_row => v_rownum,p_string => cell_xml.cell_text,p_clob => v_clob);
              END IF;              
            exception
              WHEN no_data_found THEN
                null;
            end;
         end loop;         
         add(v_clob,'</row>'||chr(10));         
         v_rownum := v_rownum + 1;
       end loop cells;
       
       DBMS_LOB.TRIM(v_agg_clob,0);
       v_agg_strings_cnt := 1;       
       <<aggregates>>       
       for row_xml in (select column_value as row_ from table (select xmlsequence(extract(rowset_xml.rowset,'ROWSET/AGGREGATE')) from dual)) loop

         for cell_xml_agg in (select rownum coll_num,
                                 extractvalue(COLUMN_VALUE, 'CELL') AS cell_text,
                                 extractvalue(COLUMN_VALUE, 'CELL/@value') AS cell_value
                          from table (select xmlsequence(extract(row_xml.row_,'AGGREGATE/CELL')) from dual))
         loop
           v_agg_strings_cnt := greatest(length(regexp_replace('[^:]','')) + 1,v_agg_strings_cnt);
           if instr(cell_xml_agg.cell_text,':') > 0 then
             print_char_cell(p_coll => cell_xml_agg.coll_num,
                             p_row => v_rownum,
                             p_string => rtrim(cell_xml_agg.cell_text,chr(10)),
                             p_clob => v_agg_clob,
                             p_style_id => aggregate_style_id);
           else
             print_number_cell(p_coll => cell_xml_agg.coll_num, 
                               p_row => v_rownum, 
                               p_value => cell_xml_agg.cell_value,
                               p_clob => v_agg_clob,
                               p_style_id => aggregate_style_id);
           end if;
         end loop;
         add(v_clob,'<row ht="'||v_agg_strings_cnt*string_height||'">'||chr(10));
         --add(v_clob,'<row ht="40">'||chr(10));
         dbms_lob.copy( dest_lob => v_clob,
                        src_lob => v_agg_clob,
                        amount => dbms_lob.getlength(v_agg_clob),
                        dest_offset => dbms_lob.getlength(v_clob),
                        src_offset => 1);
         add(v_clob,'</row>'||chr(10));
         v_rownum := v_rownum + 1;
       end loop aggregates;   
     end loop rowset;     
     
     add(v_clob,'</sheetData>'||chr(10));
     --if p_autofilter then
     add(v_clob,'<autoFilter ref="A1:' || get_cell_name(v_colls_count,v_rownum) || '"/>');
     --end if;
     add(v_clob,'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'||chr(10));
     
     add(v_strings_clob,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'||chr(10));
     add(v_strings_clob,'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' || v_strings.count() || '" uniqueCount="' || v_strings.count() || '">'||chr(10));
        
     for i in 1 .. v_strings.count() loop
        add(v_strings_clob,'<si><t>'||dbms_xmlgen.convert( substr( v_strings( i ), 1, 32000 ) ) || '</t></si>'||chr(10));
     end loop; 
     add(v_strings_clob,'</sst>'||chr(10));
     
     dbms_lob.freetemporary(v_agg_clob);
  end get_excel;
  ------------------------------------------------------------------------------
  procedure download_file(p_app_id      in number,
                          p_page_id     in number,
                          p_max_rows    in number,
                          p_file_name   in varchar2 default 'Excel')
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
    
    
    get_excel(v_xml_data,v_cells,v_strings);
    add1file( t_excel, 'xl/worksheets/Sheet1.xml', v_cells);
    add1file( t_excel, 'xl/sharedStrings.xml',v_strings);
    add1file( t_excel, 'xl/_rels/workbook.xml.rels',t_sheet_rels);    
    add1file( t_excel, 'xl/workbook.xml',t_workbook);    
    
    as_zip.finish_zip( t_excel );
      
    htp.flush;
    owa_util.mime_header( wwv_flow_utilities.get_excel_mime_type, false );
    htp.print( 'Content-Length: ' || dbms_lob.getlength( t_excel ) );
    htp.print( 'Content-disposition: attachment; filename='||p_file_name||'.xlsx;' );
    owa_util.http_header_close;
    wpg_docload.download_file( t_excel );
    dbms_lob.freetemporary(t_excel);
    dbms_lob.freetemporary(v_cells);
    dbms_lob.freetemporary(v_strings);    
  end download_file;
  
end;
/
