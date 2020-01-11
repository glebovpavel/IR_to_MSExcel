CREATE OR REPLACE PACKAGE ir_to_xlsx
  AUTHID current_user
IS
  PROCEDURE download_excel(p_app_id       IN NUMBER,
                           p_page_id      IN NUMBER,
                           p_region_id    IN NUMBER,
                           p_col_length   IN VARCHAR2 DEFAULT NULL,
                           p_max_rows     IN NUMBER,
                           p_autofilter   IN CHAR DEFAULT 'Y',
                           p_export_links IN CHAR DEFAULT 'N',
                           p_custom_width IN VARCHAR2
                          ); 
  
  FUNCTION convert_date_format(p_format IN VARCHAR2)
  RETURN VARCHAR2;
  
  FUNCTION convert_number_format(p_format IN VARCHAR2)
  RETURN VARCHAR2;  
  
  FUNCTION convert_date_format(p_datatype IN VARCHAR2,p_format IN VARCHAR2)
  RETURN VARCHAR2;
  
  FUNCTION convert_date_format_js(p_datatype IN VARCHAR2, p_format IN VARCHAR2)
  RETURN VARCHAR2;
  
  FUNCTION get_max_rows (p_app_id      IN NUMBER,
                         p_page_id     IN NUMBER,
                         p_region_id   IN NUMBER)
  RETURN NUMBER;

  FUNCTION get_highlight_in_cond_sql(p_condition_expression  IN apex_application_page_ir_cond.condition_expression%TYPE,
                                     p_condition_sql         IN apex_application_page_ir_cond.condition_sql%TYPE,
                                     p_condition_column_name IN apex_application_page_ir_cond.condition_column_name%TYPE)
  RETURN VARCHAR2;
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
                    
END;
/

