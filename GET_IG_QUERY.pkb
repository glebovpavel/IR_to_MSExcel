create or replace package body GET_IG_QUERY
as
    FUNCTION ir_query_where(app_id_in IN NUMBER,
                            page_id_in IN NUMBER,
                            session_id_in IN NUMBER,
                            base_report_id_in IN VARCHAR2) RETURN VARCHAR2 IS
        /* 
            Parameters:     base_report_id_in - User's currently-displayed report (including saved)
        
            Returns:        ANDed WHERE clause to be run against base view
        
            Author:         STEWART_L_STRYKER
            Created:        2/12/2009 5:16:51 PM
        
            Usage:          RETURN apex_ir_query.ir_query_where(app_id_in => :APP_ID, 
                                        page_id_in => 2, 
                                        session_id_in => :APP_SESSION);
        
            CS-RCS Modification History: (Do NOT edit manually)
        
            $Log: $
        */
        query_string VARCHAR2(32500);
        test_val VARCHAR2(80);
        search_count PLS_INTEGER := 0;
        clause VARCHAR2(32000);
        
        query_too_long EXCEPTION;
        PRAGMA EXCEPTION_INIT(query_too_long, -24381);
    BEGIN
        FOR filters IN (SELECT condition_column_name,
                               condition_operator,
                               condition_expression,
                               condition_expression2
                          FROM apex_application_page_ir_cond cond
                          JOIN apex_application_page_ir_rpt r ON r.application_id =
                                                                 cond.application_id
                                                             AND r.page_id = cond.page_id
                                                             AND r.report_id = cond.report_id
                         WHERE cond.application_id = app_id_in
                           AND cond.page_id = page_id_in
                           AND cond.condition_type = 'Filter'
                           AND cond.condition_enabled = 'Yes'
                           AND r.base_report_id = base_report_id_in
                           AND r.session_id = session_id_in)
        LOOP
            clause := ir_query_parse_filter(filters.condition_column_name,
                                            filters.condition_operator,
                                            filters.condition_expression,
                                            filters.condition_expression2);
            IF LENGTH(clause) + LENGTH(query_string) > 32500
            THEN
                RAISE query_too_long;
            END IF;
            
            query_string := query_string || ' AND ' || clause;                            
        END LOOP;
    

        FOR searches IN (SELECT r.report_columns,
                                cond.condition_expression,
                                to_char(r.interactive_report_id) AS interactive_report_id
                          FROM apex_application_page_ir_cond cond
                          JOIN apex_application_page_ir_rpt r ON r.application_id =
                                                                 cond.application_id
                                                             AND r.page_id = cond.page_id
                                                             AND r.report_id = cond.report_id
                         WHERE cond.application_id = app_id_in
                           AND cond.page_id = page_id_in
                           AND cond.condition_type = 'Search'
                           AND cond.condition_enabled = 'Yes'
                           AND r.base_report_id = base_report_id_in
                           AND r.session_id = session_id_in)
        LOOP
            search_count := search_count + 1;
            test_val := NVL(searches.interactive_report_id, 'null');
            clause := ir_query_parse_search(searches.report_columns,
                                            searches.condition_expression,
                                            app_id_in,
                                            searches.interactive_report_id);

            IF LENGTH(clause) + LENGTH(query_string) > 32500
            THEN
                RAISE query_too_long;
            END IF;
            
            query_string := query_string || ' AND ' || clause;
                            
        END LOOP;

        log_apex_access(app_name_in => app_id_in,
                                          app_user_in => v('APP_USER'),
                                          msg_in      => 'Searches: ' || search_count ||
                                                         '.  base_report_id_in: ' || nvl(base_report_id_in, 'null')
                                                         || '.  Session: ' || session_id_in);
        RETURN query_string;
    EXCEPTION
        WHEN query_too_long THEN
            log_apex_access(app_name_in => app_id_in,
                                              app_user_in => v('APP_USER'),
                                              msg_in      => 'Generated query string would have been > 32k');
        
            RETURN query_string;
        WHEN no_data_found THEN
            log_apex_access(app_name_in => app_id_in,
                                              app_user_in => v('APP_USER'),
                                              msg_in      => 'NDF. Searches: ' || search_count ||
                                                             '.  IR Report id: ' || nvl(test_val, 'null'));
        
            RETURN query_string;
        WHEN OTHERS THEN
            log_apex_access(app_name_in => app_id_in,
                                              app_user_in => v('APP_USER'),
                                              msg_in      => 'EXCEPTION: ' || SQLERRM ||
                                                             '.  Searches: ' || search_count ||
                                                             '.  IR Report id: ' || nvl(test_val, 'null'));
        
            RETURN query_string;
    END ir_query_where;


end GET_IG_QUERY; ﻿