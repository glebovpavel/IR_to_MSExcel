CREATE OR REPLACE package IR_TO_MSEXCEL 
as
  function get_xlsx_from_ir (p_process in apex_plugin.t_process,
                                 p_plugin  in apex_plugin.t_plugin )
  return apex_plugin.t_process_exec_result;
   
  procedure get_xlsx_from_ir_ext(p_maximum_rows    in number default null,
                                 p_jquery_selector in varchar2 default null,
                                 p_download_type   in char default 'E',   -- E -> Excel XLSX, X -> XML (Debug), T -> Debug TXT
                                 p_replace_xls     in char default 'Y',   --Y/N
                                 p_custom_width    in varchar2 default null
                                );
  -- p_custom_width is delimetered string of COLUMN_NAME,COLUMN_WIDTH=COLUMN_NAME,COLUMN_WIDTH=  etc.
  -- sample: PROJECT,151=TASK_NAME,319=START_DATE,133=                    
                                
end IR_TO_MSEXCEL;
/


CREATE OR REPLACE package body IR_TO_MSEXCEL 
as
    XLS_DOWNLOAD_SELECTOR constant varchar2(20) := ' #apexir_dl_XLS, ';

    JAVASCRIPT_CODE  constant varchar2(850) := 
    q'[   
    function getColWidthsDelimeteredString()
      {
        var colWidthsDelimeteredString = "";
        var colWidthsArray = Array ();        
        $( ".apexir_WORKSHEET_DATA th" ).each(function( index,elmt ) 
        {    
          colWidthsArray[elmt.id] = $(elmt).width();  
        });
        for (var i in colWidthsArray) {
         colWidthsDelimeteredString = colWidthsDelimeteredString + i + '=' + colWidthsArray[i] + "\,";  
        }       
        return colWidthsDelimeteredString + '#CUSTOMWIDTH#';
      }
    function replaceDownloadXLS()
      {
        $("#apexir_CONTROL_PANEL_DROP").on('click',function(){                
           $("#apexir_dl_XLS").attr("href",'f?p=&APP_ID.:&APP_PAGE_ID.:&APP_SESSION.:GPV_IR_TO_MSEXCEL' + getColWidthsDelimeteredString() +':NO:::');
        });
      }
    ]';
   ------------------------------------------------------------------------------ 
    ON_SELECOR_CODE constant varchar2(220) :=  q'[
    $("#SELECTOR#").on( "click", function() 
    { 
      apex.navigation.redirect("f?p=&APP_ID.:&APP_PAGE_ID.:&APP_SESSION.:GPV_IR_TO_MSEXCEL"+getColWidthsDelimeteredString()+":NO:::") ;
    });]';
  ------------------------------------------------------------------------------  
    STANDARD_DOWNLOAD_CODE constant varchar2(150) := q'[
    $("#apexir_WORKSHEET_REGION").bind("apexafterrefresh", function(){        
        replaceDownloadXLS()  
    });
    replaceDownloadXLS();
    ]';
  ------------------------------------------------------------------------------ 
  
  procedure get_xlsx_from_ir_ext(p_maximum_rows    in number default null,
                                 p_jquery_selector in varchar2 default null,
                                 p_download_type   in char default 'E',   -- E -> Excel XLSX, X -> XML (Debug), T -> Debug TXT
                                 p_replace_xls     in char default 'Y',   --Y/N
                                 p_custom_width    in varchar2     
                                )
  is
    v_javascript_code       varchar2(1500);
    v_xls_download_selector varchar2(500);
  begin
    v_javascript_code := JAVASCRIPT_CODE;    
   
    if p_replace_xls = 'Y' then
      v_javascript_code := v_javascript_code||STANDARD_DOWNLOAD_CODE;
      v_xls_download_selector := XLS_DOWNLOAD_SELECTOR;
    end if;

    if p_jquery_selector is not null then
     v_javascript_code := v_javascript_code||replace(ON_SELECOR_CODE,'#SELECTOR#',rtrim(v_xls_download_selector||p_jquery_selector,','));
    end if;
    
    v_javascript_code := replace(v_javascript_code,'#CUSTOMWIDTH#',replace(p_custom_width,'''',''));
    
    
    APEX_JAVASCRIPT.ADD_ONLOAD_CODE(apex_plugin_util.replace_substitutions(v_javascript_code));
  
    if v('REQUEST') like 'GPV_IR_TO_MSEXCEL%' then
      if p_download_type = 'E' then -- Excel XLSX
        XML_TO_XSLX.download_file(p_app_id       => v('APP_ID'),
                                  p_page_id      => v('APP_PAGE_ID'),
                                  p_col_length   => regexp_replace(v('REQUEST'),'^GPV_IR_TO_MSEXCEL',''),
                                  p_max_rows     => nvl(p_maximum_rows,xml_to_xslx.get_max_rows (v('APP_ID'),v('APP_PAGE_ID')))
                                  );
      elsif p_download_type = 'X' then -- XML
        IR_TO_XML.get_report_xml(p_app_id            => v('APP_ID'),
                                 p_page_id           => v('APP_PAGE_ID'),       
                                 p_return_type       => 'X',                        
                                 p_get_page_items    => 'N',
                                 p_items_list        => null,
                                 p_collection_name   => null,
                                 p_max_rows          => xml_to_xslx.get_max_rows (v('APP_ID'),v('APP_PAGE_ID'))
                                );
      elsif p_download_type = 'T' then -- Debug txt
        IR_TO_XML.get_report_xml(p_app_id            => v('APP_ID'),
                                 p_page_id           => v('APP_PAGE_ID'),       
                                 p_return_type       => 'Q',                        
                                 p_get_page_items    => 'N',
                                 p_items_list        => null,
                                 p_collection_name   => null,
                                 p_max_rows          => xml_to_xslx.get_max_rows (v('APP_ID'),v('APP_PAGE_ID'))
                                );
    
     else
      raise_application_error(-20001,'GPV_IR_TO_MSEXCEL : unknown Return Type');
     end if;
   end if;
   
  end get_xlsx_from_ir_ext;
  ------------------------------------------------------------------------------ 
  procedure check_correct_use
  is
     v_process_name  apex_application_page_proc.process_name%TYPE;
     v_page_id       apex_application_page_proc.page_id%TYPE;
  begin
    select process_name, 
           page_id
      into v_process_name,
           v_page_id
      from apex_application_page_proc 
      where application_id = 102
        and process_type_code = 'PLUGIN_GPV_IR_TO_MSEXCEL'
        and process_point_code != 'BEFORE_BOX_BODY';
  
       raise_application_error(-20001,'Plugin "GPV Interactive Report to MSExcel" must be used in "On Load - Before Region" processes only. Please check page '||v_page_id||'.');  
  exception
    when no_data_found then
       null; --check ok
  end check_correct_use;
  ------------------------------------------------------------------------------ 
  FUNCTION get_xlsx_from_ir (p_process IN apex_plugin.t_process,
                             p_plugin  IN apex_plugin.t_plugin )
  RETURN apex_plugin.t_process_exec_result  
  is
    v_javascript_code varchar2(2000);
    v_on_standard_download_code varchar2(300);
    v_on_selecor_code varchar2(300);
  BEGIN
    check_correct_use;
    get_xlsx_from_ir_ext(p_maximum_rows    => p_process.attribute_05,
                         p_jquery_selector => p_process.attribute_06,
                         p_download_type   => p_process.attribute_07,
                         p_replace_xls     => p_process.attribute_10
                        );
   
   return null;
  end get_xlsx_from_ir;

end IR_TO_MSEXCEL;
/
