create or replace PACKAGE  "IR_TO_MSEXCEL" 
  AUTHID CURRENT_USER
as
  FUNCTION render  (p_dynamic_action in apex_plugin.t_dynamic_action,
                    p_plugin         in apex_plugin.t_plugin )
  return apex_plugin.t_dynamic_action_render_result; 
  
  function ajax (p_dynamic_action in apex_plugin.t_dynamic_action,
                 p_plugin         in apex_plugin.t_plugin )
  return apex_plugin.t_dynamic_action_ajax_result;
                                
end IR_TO_MSEXCEL;
/

create or replace PACKAGE BODY  "IR_TO_MSEXCEL" 
as
  
  function get_affected_region_id(p_dynamic_action_id IN apex_application_page_da_acts.action_id%TYPE,
                                  p_html_region_id    IN VARCHAR2
                                 )
  return  apex_application_page_da_acts.affected_region_id%type
  is
   v_affected_region_id apex_application_page_da_acts.affected_region_id%type;
  begin
  
      SELECT affected_region_id
      INTO v_affected_region_id
      FROM apex_application_page_da_acts aapda
      WHERE aapda.action_id = p_dynamic_action_id
        and page_id in(v('APP_PAGE_ID'),0)
        and application_id = v('APP_ID')
        and rownum <2; 
      
      if v_affected_region_id is null then 
        begin 
          select region_id 
          into v_affected_region_id 
          from apex_application_page_regions 
          where  static_id = p_html_region_id
            and page_id in (v('APP_PAGE_ID'),0)
            and application_id = v('APP_ID');
        exception 
          when no_data_found then 
           select region_id 
           into v_affected_region_id 
           from apex_application_page_regions 
           where  region_id = ltrim(p_html_region_id,'R')
             and page_id in (v('APP_PAGE_ID'),0)
             and application_id = v('APP_ID'); 
        end; 
      end if;       
      return v_affected_region_id;
  exception
    when others then
      raise_application_error(-20001,'IR_TO_MSEXCEL.get_affected_region_id: No region found!');
  end get_affected_region_id;
  ------------------------------------------------------------------------------
  
  function get_affected_region_static_id(p_dynamic_action_id IN apex_application_page_da_acts.action_id%TYPE)
  return  apex_application_page_regions.static_id%TYPE
  is
   v_affected_region_selector apex_application_page_regions.static_id%type;
  begin
      SELECT nvl(static_id,'R'||to_char(affected_region_id)) 
      INTO v_affected_region_selector 
      FROM apex_application_page_da_acts aapda, 
           apex_application_page_regions r 
      WHERE aapda.action_id = p_dynamic_action_id 
        and aapda.affected_region_id = r.region_id 
        and r.source_type ='Interactive Report' 
        and aapda.page_id = v('APP_PAGE_ID') 
        and aapda.application_id = v('APP_ID')
        and r.page_id = v('APP_PAGE_ID') 
        and r.application_id = v('APP_ID'); 
      
      return v_affected_region_selector;
  exception
    when no_data_found then  
      return null;
  end get_affected_region_static_id;
  ------------------------------------------------------------------------------
  FUNCTION get_version
  return varchar2
  is
   v_version varchar2(3);
  begin
     SELECT substr(version_no,1,3) 
     into v_version
     FROM apex_release
     where rownum <2;

     return v_version;
  end get_version;

  ------------------------------------------------------------------------------
  FUNCTION render (p_dynamic_action in apex_plugin.t_dynamic_action,
                   p_plugin         in apex_plugin.t_plugin )
  return apex_plugin.t_dynamic_action_render_result
  is
    v_javascript_code          varchar2(1000);
    v_result                   apex_plugin.t_dynamic_action_render_result;
    v_plugin_id                varchar2(100);
    v_affected_region_selector apex_application_page_regions.static_id%type;
  BEGIN
    v_plugin_id := apex_plugin.get_ajax_identifier;
    v_affected_region_selector := get_affected_region_static_id(p_dynamic_action.ID);
    if nvl(p_dynamic_action.attribute_03,'Y') = 'Y' then
      if v_affected_region_selector is not null then 
        -- add XLSX Icon to Affected IR Region
        v_javascript_code :=  'excel_gpv.addDownloadXLSXIcon('''||v_plugin_id||''','''||v_affected_region_selector||''','''||get_version||''');';
        APEX_JAVASCRIPT.ADD_ONLOAD_CODE(v_javascript_code,v_affected_region_selector);
      else
        -- add XLSX Icon to all IR Regions on the page
        for i in (SELECT nvl(static_id,'R'||to_char(region_id)) as affected_region_selector      
                  FROM apex_application_page_regions r
                  where r.page_id = v('APP_PAGE_ID')
                    and r.application_id =v('APP_ID')
                    and r.source_type ='Interactive Report'
                 )
        loop         
           v_javascript_code :=  'excel_gpv.addDownloadXLSXIcon('''||v_plugin_id||''','''||i.affected_region_selector||''','''||get_version||''');';
           APEX_JAVASCRIPT.ADD_ONLOAD_CODE(v_javascript_code,i.affected_region_selector);     
        end loop;
      end if;
    end if;
   
      
    apex_javascript.add_library (p_name      => 'IR2MSEXCEL', 
                                 p_directory => p_plugin.file_prefix); 
    
    if v_affected_region_selector is not null then
      v_result.javascript_function := 'function(){excel_gpv.getExcel('''||v_affected_region_selector||''','''||v_plugin_id||''')}';
    else
     v_result.javascript_function := 'function(){console.log("No Affected Region Found!");}';
    end if;
    v_result.ajax_identifier := v_plugin_id;
    
    return v_result;
  end render;
  ------------------------------------------------------------------------------
  
  function ajax (p_dynamic_action in apex_plugin.t_dynamic_action,
                 p_plugin         in apex_plugin.t_plugin )
  return apex_plugin.t_dynamic_action_ajax_result
  is
    p_download_type      varchar2(1);
    p_custom_width       varchar2(1000);
	p_autofilter              char;
    v_maximum_rows       number;
    v_dummy              apex_plugin.t_dynamic_action_ajax_result;
    v_affected_region_id apex_application_page_da_acts.affected_region_id%type;
  begin      
      p_download_type:= nvl(p_dynamic_action.attribute_02,'E');
	  p_autofilter:= nvl(p_dynamic_action.attribute_04,'Y');
      v_affected_region_id := get_affected_region_id(p_dynamic_action_id => p_dynamic_action.ID
                                                    ,p_html_region_id    => apex_application.g_x03);
      
      v_maximum_rows := nvl(nvl(p_dynamic_action.attribute_01,
                                IR_TO_XLSX.get_max_rows (p_app_id    => apex_application.g_x01,
                                                          p_page_id   => apex_application.g_x02,
                                                          p_region_id => v_affected_region_id)
                                ),1000);                                               
      if p_download_type = 'E' then -- Excel XLSX
        IR_TO_XLSX.download_file(p_app_id       => apex_application.g_x01,
                                  p_page_id      => apex_application.g_x02,
                                  p_region_id    => v_affected_region_id,
                                  p_col_length   => apex_application.g_x04||p_custom_width,
                                  p_max_rows     => v_maximum_rows,
                                  p_autofilter => p_autofilter
                                  );
      elsif p_download_type = 'X' then -- XML
        IR_TO_XML.get_report_xml(p_app_id            => apex_application.g_x01,
                                 p_page_id           => apex_application.g_x02, 
                                 p_region_id         => v_affected_region_id,
                                 p_return_type       => 'X',                        
                                 p_get_page_items    => 'N',
                                 p_items_list        => null,
                                 p_collection_name   => null,
                                 p_max_rows          => v_maximum_rows
                                );
      elsif p_download_type = 'T' then -- Debug txt
        IR_TO_XML.get_report_xml(p_app_id            => apex_application.g_x01,
                                 p_page_id           => apex_application.g_x02, 
                                 p_region_id         => v_affected_region_id,
                                 p_return_type       => 'Q',                        
                                 p_get_page_items    => 'N',
                                 p_items_list        => null,
                                 p_collection_name   => null,
                                 p_max_rows          => v_maximum_rows
                                );
      elsif p_download_type = 'M' then -- use Moritz Klein engine https://github.com/commi235                        
       apexir_xlsx_pkg.download(  p_ir_region_id   => v_affected_region_id,
                                  p_app_id         => apex_application.g_x01,
                                  p_ir_page_id     => apex_application.g_x02
                                );
     else
      raise_application_error(-20001,'GPV_IR_TO_MSEXCEL : unknown Return Type');
     end if;    
     return v_dummy;
  exception
    when others then
      raise_application_error(-20001,SQLERRM||chr(10)||dbms_utility.format_error_backtrace);
  end ajax;
  
end IR_TO_MSEXCEL;
/

