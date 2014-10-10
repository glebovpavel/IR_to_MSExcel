CREATE OR REPLACE package GPV_IR_TO_MSEXCEL 
as
  function gpv_get_xlsx_from_ir (p_process in apex_plugin.t_process,
                                 p_plugin  in apex_plugin.t_plugin )
   return apex_plugin.t_process_exec_result;
end GPV_IR_TO_MSEXCEL;
/


CREATE OR REPLACE package body GPV_IR_TO_MSEXCEL 
as
  FUNCTION gpv_get_xlsx_from_ir (p_process IN apex_plugin.t_process,
                                 p_plugin  IN apex_plugin.t_plugin )
  RETURN apex_plugin.t_process_exec_result  
  is
    v_javascript_code varchar2(2000);
    v_on_standard_download_code varchar2(300);
    v_on_selecor_code varchar2(300);
  BEGIN
    --FULL VERSION
    
    v_javascript_code := 
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
         colWidthsDelimeteredString = colWidthsDelimeteredString + i + '\,' + colWidthsArray[i] + "=";  
        }       
        return colWidthsDelimeteredString;
      }
    ]';
    
    v_on_selecor_code :=  q'[
    $("#SELECTOR#").on( "click", function() 
    { 
      apex.navigation.redirect("f?p=&APP_ID.:&APP_PAGE_ID.:&APP_SESSION.:GPV_IR_TO_MSEXCEL"+getColWidthsDelimeteredString()+":NO:::") ;
    });]';
    
    v_on_standard_download_code := q'[
    $("#apexir_CONTROL_PANEL_DROP").on('click',function(){                
       $("#apexir_dl_XLS").attr("href",'f?p=&APP_ID.:&APP_PAGE_ID.:&APP_SESSION.:GPV_IR_TO_MSEXCEL' + getColWidthsDelimeteredString() +':NO:::');
    });   ]';
    
    if p_process.attribute_06 is not null then
     v_javascript_code := v_javascript_code||replace(v_on_selecor_code,'#SELECTOR#',p_process.attribute_06);
    end if;
    
    if p_process.attribute_10 = 'Y' then
      v_javascript_code := v_javascript_code||v_on_standard_download_code;
    end if;
    
    
    --MIN VERSION
    --v_javascript_code := q'[$("'||p_process.attribute_05||'").on("click",function(){var i="",_=Array();$(".apexir_WORKSHEET_DATA th").each(function(i,a){_[a.id]=$(a).width()});for(var a in _)i=i+a+","+_[a]+"=";apex.navigation.redirect("f?p=&APP_ID.:&APP_PAGE_ID.:&APP_SESSION.:GPV_IR_TO_MSEXCEL"+i+":NO:::")});]';
    
    APEX_JAVASCRIPT.ADD_ONLOAD_CODE(apex_plugin_util.replace_substitutions(v_javascript_code));
  
    if v('REQUEST') like 'GPV_IR_TO_MSEXCEL%' then
      if p_process.attribute_07 = 'E' then -- Excel XLSX
        XML_TO_XSLX.download_file(p_app_id       => v('APP_ID'),
                                  p_page_id      => v('APP_PAGE_ID'),
                                  p_max_rows     => p_process.attribute_05,
                                  p_col_length   => regexp_replace(v('REQUEST'),'^GPV_IR_TO_MSEXCEL',''),
                                  p_coefficient  => p_process.attribute_09
                                  );
      elsif p_process.attribute_07 = 'X' then -- XML
        IR_TO_XML.get_report_xml(p_app_id            => v('APP_ID'),
                                 p_page_id           => v('APP_PAGE_ID'),       
                                 p_return_type       => 'X',                        
                                 p_get_page_items    => 'N',
                                 p_items_list        => null,
                                 p_collection_name   => null,
                                 p_max_rows          => p_process.attribute_05
                                );
      elsif p_process.attribute_07 = 'T' then -- Debug txt
        IR_TO_XML.get_report_xml(p_app_id            => v('APP_ID'),
                                 p_page_id           => v('APP_PAGE_ID'),       
                                 p_return_type       => 'Q',                        
                                 p_get_page_items    => 'N',
                                 p_items_list        => null,
                                 p_collection_name   => null,
                                 p_max_rows          => p_process.attribute_05
                                );
    
     else
      raise_application_error(-20001,'GPV_IR_TO_MSEXCEL : unknown Return Type');
     end if;
   end if;
   
   return null;
  end gpv_get_xlsx_from_ir;

end GPV_IR_TO_MSEXCEL;
/
