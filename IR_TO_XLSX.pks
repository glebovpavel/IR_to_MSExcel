CREATE OR REPLACE PACKAGE IR_TO_XLSX
  AUTHID CURRENT_USER
IS
  procedure download_excel(p_app_id       IN NUMBER,
                           p_page_id      IN NUMBER,
                           p_region_id    IN NUMBER,
                           p_col_length   IN VARCHAR2 DEFAULT NULL,
                           p_max_rows     IN NUMBER,
                           p_autofilter   IN CHAR DEFAULT 'Y'
                          ); 
  procedure download_debug(p_app_id      IN NUMBER,
                           p_page_id     IN NUMBER,
                           p_region_id   IN NUMBER, 
                           p_col_length  IN VARCHAR2 DEFAULT NULL,
                           p_max_rows    IN NUMBER,
                           p_autofilter  IN CHAR DEFAULT 'Y'                           
                          );
                          
  function convert_date_format(p_format IN VARCHAR2)
  return varchar2;
  function convert_number_format(p_format IN VARCHAR2)
  return varchar2;
  
  function get_max_rows (p_app_id      IN NUMBER,
                         p_page_id     IN NUMBER,
                         p_region_id   IN NUMBER)
  return number;

   /* 
    function to handle cases of 'in' and 'not in' conditions for highlights
       used in cursor cur_highlight
    
    Author: Srihari Ravva
  */ 
  function get_highlight_in_cond_sql(p_condition_expression  in APEX_APPLICATION_PAGE_IR_COND.CONDITION_EXPRESSION%TYPE,
                                     p_condition_sql         in APEX_APPLICATION_PAGE_IR_COND.CONDITION_SQL%TYPE,
                                     p_condition_column_name in APEX_APPLICATION_PAGE_IR_COND.CONDITION_COLUMN_NAME%TYPE)
  return varchar2;
  /*
  -- format test cases
  select ir_to_xlsx.convert_date_format('dd.mm.yyyy hh24:mi:ss'),to_char(sysdate,'dd.mm.yyyy hh24:mi:ss') from dual
  union
  select ir_to_xlsx.convert_date_format('dd.mm.yyyy hh12:mi:ss'),to_char(sysdate,'dd.mm.yyyy hh12:mi:ss') from dual
  union
  select ir_to_xlsx.convert_date_format('day-mon-yyyy'),to_char(sysdate,'day-mon-yyyy') from dual
  union
  select ir_to_xlsx.convert_date_format('month'),to_char(sysdate,'month') from dual
  union
  select ir_to_xlsx.convert_date_format('RR-MON-DD'),to_char(sysdate,'RR-MON-DD') from dual 
  union
  select ir_to_xlsx.convert_number_format('FML999G999G999G999G990D0099'),to_char(123456789/451,'FML999G999G999G999G990D0099') from dual
  union
  select ir_to_xlsx.convert_date_format('DD-MON-YYYY HH:MIPM'),to_char(sysdate,'DD-MON-YYYY HH:MIPM') from dual 
  union
  select ir_to_xlsx.convert_date_format('fmDay, fmDD fmMonth, YYYY'),to_char(sysdate,'fmDay, fmDD fmMonth, YYYY') from dual 
  */
                    
end;
/

