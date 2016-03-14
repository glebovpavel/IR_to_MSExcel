CREATE OR REPLACE PACKAGE "APEXIR_XLSX_PKG" 
  AUTHID CURRENT_USER
AS
  /*
   This is mock package to make other packages compilable
   
   to use original commi235/APEX_IR_XLSX generating engine,
   please download it from https://github.com/commi235/APEX_IR_XLSX

  */

  PROCEDURE download
    ( p_ir_region_id NUMBER := NULL
    , p_app_id NUMBER := NV('APP_ID')
    , p_ir_page_id NUMBER := NV('APP_PAGE_ID')
    , p_ir_session_id NUMBER := NV('SESSION')
    , p_ir_request VARCHAR2 := V('REQUEST')
    , p_ir_view_mode VARCHAR2 := NULL
    , p_column_headers BOOLEAN := TRUE
    , p_col_hdr_help BOOLEAN := TRUE
    , p_aggregates IN BOOLEAN := TRUE
    , p_process_highlights IN BOOLEAN := TRUE
    , p_show_report_title IN BOOLEAN := TRUE
    , p_show_filters IN BOOLEAN := TRUE
    , p_show_highlights IN BOOLEAN := TRUE
    , p_original_line_break IN VARCHAR2 := '<br />'
    , p_replace_line_break IN VARCHAR2 := chr(13) || chr(10)
    , p_append_date IN BOOLEAN := TRUE
    );
END APEXIR_XLSX_PKG;
/


CREATE OR REPLACE PACKAGE BODY "APEXIR_XLSX_PKG" 
AS

  /*
   This is mock package to make other packages compilable
   
   to use original commi235/APEX_IR_XLSX generating engine,
   please download it from https://github.com/commi235/APEX_IR_XLSX

  */

  PROCEDURE download
    ( p_ir_region_id NUMBER := NULL
    , p_app_id NUMBER := NV('APP_ID')
    , p_ir_page_id NUMBER := NV('APP_PAGE_ID')
    , p_ir_session_id NUMBER := NV('SESSION')
    , p_ir_request VARCHAR2 := V('REQUEST')
    , p_ir_view_mode VARCHAR2 := NULL
    , p_column_headers BOOLEAN := TRUE
    , p_col_hdr_help BOOLEAN := TRUE
    , p_aggregates IN BOOLEAN := TRUE
    , p_process_highlights IN BOOLEAN := TRUE
    , p_show_report_title IN BOOLEAN := TRUE
    , p_show_filters IN BOOLEAN := TRUE
    , p_show_highlights IN BOOLEAN := TRUE
    , p_original_line_break IN VARCHAR2 := '<br />'
    , p_replace_line_break IN VARCHAR2 := chr(13) || chr(10)
    , p_append_date IN BOOLEAN := TRUE
    )
  AS
  BEGIN
   raise_application_error(-20001,'commi235 generating engine not installed! Please read documentation!');
  END download;

END APEXIR_XLSX_PKG;
/
