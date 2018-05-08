create or replace PACKAGE BODY IR_TO_MSEXCEL 
as
  subtype t_large_varchar2  is varchar2(32767);
  v_plugin_running boolean default false;
  
  function is_ir2msexcel 
  return boolean
  is
  begin
    return nvl(v_plugin_running,false);    
  end is_ir2msexcel;
  ------------------------------------------------------------------------------
  
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
        and page_id in(nv('APP_PAGE_ID'),0)
        and application_id = nv('APP_ID')
        and rownum <2; 
      
      if v_affected_region_id is null then 
        begin 
          select region_id 
          into v_affected_region_id 
          from apex_application_page_regions 
          where  static_id = p_html_region_id
            and page_id in (nv('APP_PAGE_ID'),0)
            and application_id = nv('APP_ID');
        exception 
          when no_data_found then 
           select region_id 
           into v_affected_region_id 
           from apex_application_page_regions 
           where  region_id = to_number(ltrim(p_html_region_id,'R'))
             and page_id in (nv('APP_PAGE_ID'),0)
             and application_id = nv('APP_ID'); 
        end; 
      end if;       
      return v_affected_region_id;
  exception
    when others then
      raise_application_error(-20001,'IR_TO_MSEXCEL.get_affected_region_id: No region found!');      
  end get_affected_region_id;
  ------------------------------------------------------------------------------
  
  function get_affected_region_static_id(p_dynamic_action_id IN apex_application_page_da_acts.action_id%TYPE,
                                         p_type              IN varchar2   
                                         )
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
        and r.source_type = p_type
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
  
  function get_ig_file_name (p_region_selector IN VARCHAR2)
  return varchar2
  is
    v_filename APEX_APPL_PAGE_IGS.download_filename%TYPE;
  begin
    select download_filename
      into v_filename
      from APEX_APPL_PAGE_IGS
    where application_id = v('APP_ID')
      and page_id = v('APP_PAGE_ID')
      and nvl(region_name,'R'||region_id) = p_region_selector
      and rownum <2;

     return apex_plugin_util.replace_substitutions(nvl(v_filename,'Excel'));
  exception
    when others then
       return 'Excel';
  end get_ig_file_name;
  ------------------------------------------------------------------------------
  
  FUNCTION render (p_dynamic_action in apex_plugin.t_dynamic_action,
                   p_plugin         in apex_plugin.t_plugin )
  return apex_plugin.t_dynamic_action_render_result
  is
    v_javascript_code          varchar2(1000);
    v_result                   apex_plugin.t_dynamic_action_render_result;
    v_plugin_id                varchar2(100);
    v_affected_region_IR_selector apex_application_page_regions.static_id%type;
    v_affected_region_IG_selector apex_application_page_regions.static_id%type;
    v_is_ig                    boolean default false;
    v_is_ir                    boolean default false;
    v_workspace                apex_applications.workspace%TYPE;
    v_found                    boolean default false;
  BEGIN
    v_plugin_id := apex_plugin.get_ajax_identifier;
    v_affected_region_IR_selector := get_affected_region_static_id(p_dynamic_action.ID,'Interactive Report');
    v_affected_region_IG_selector := get_affected_region_static_id(p_dynamic_action.ID,'Interactive Grid');
    
    select workspace
    into v_workspace
    from apex_applications
    where application_id =nv('APP_ID');
    
    if nvl(p_dynamic_action.attribute_03,'Y') = 'Y' then -- add "Download XLSX" icon
      if v_affected_region_IR_selector is not null then
        -- add XLSX Icon to Affected IR Region
        v_javascript_code :=  'excel_gpv.addDownloadXLSXIcon('''||v_plugin_id||''','''||v_affected_region_IR_selector||''','''||get_version||''');';
        APEX_JAVASCRIPT.ADD_ONLOAD_CODE(v_javascript_code,v_affected_region_IR_selector);
      elsif v_affected_region_IG_selector is not null then
         v_javascript_code := 'excel_ig_gpv.addDownloadXLSXiconToIG('''||v_affected_region_IG_selector
           ||''','''||v_plugin_id||''','''||get_ig_file_name(v_affected_region_IG_selector)||''',''/'||ltrim(p_plugin.file_prefix,'/')||''');';  
         APEX_JAVASCRIPT.ADD_ONLOAD_CODE(v_javascript_code,v_affected_region_IG_selector);
      else
        -- add XLSX Icon to all IR Regions on the page
        for i in (SELECT nvl(static_id,'R'||to_char(region_id)) as affected_region_selector,
                         r.source_type
                  FROM apex_application_page_regions r
                  where r.page_id = nv('APP_PAGE_ID')
                    and r.application_id =nv('APP_ID')
                    and r.source_type  in ('Interactive Report','Interactive Grid')
                    and r.workspace = v_workspace
                 )
        loop
           if i.source_type = 'Interactive Report' then 
             v_javascript_code :=  'excel_gpv.addDownloadXLSXIcon('''||v_plugin_id||''','''||i.affected_region_selector||''','''||get_version||''');';
             v_is_ir := true;
           else             
             v_javascript_code := 'excel_ig_gpv.addDownloadXLSXiconToIG('''||i.affected_region_selector||''','''||v_plugin_id||''','''||get_ig_file_name(v_affected_region_IG_selector)||''',''/'||ltrim(p_plugin.file_prefix,'/')||''');';  
             v_is_ig := true;
           end if;
           APEX_JAVASCRIPT.ADD_ONLOAD_CODE(v_javascript_code,i.affected_region_selector);
        end loop;
      end if;
    end if;
    
    if v_affected_region_IR_selector is not null or v_is_ir then
       apex_javascript.add_library (p_name      => 'IR2MSEXCEL', 
                                    p_directory => p_plugin.file_prefix); 
    end if;                                 
    if v_affected_region_IG_selector is not null or v_is_ig then
        apex_javascript.add_library (p_name      => 'IG2MSEXCEL', 
                                     p_directory => p_plugin.file_prefix); 
        apex_javascript.add_library (p_name      => 'shim.min', 
                                     p_directory => p_plugin.file_prefix); 
        apex_javascript.add_library (p_name      => 'blob.min', 
                                     p_directory => p_plugin.file_prefix);
        apex_javascript.add_library (p_name      => 'FileSaver.min', 
                                     p_directory => p_plugin.file_prefix);
    end if;
    
    if v_affected_region_IR_selector is not null then
      v_result.javascript_function := 'function(){excel_gpv.getExcel('''||v_affected_region_IR_selector||''','''||v_plugin_id||''')}';
    elsif v_affected_region_IG_selector is not null then      
      v_result.javascript_function := 'function(){excel_ig_gpv.downloadXLSXfromIG('''||v_affected_region_IG_selector||''','''||v_plugin_id||''','''||get_ig_file_name(v_affected_region_IG_selector)||''',''/'||ltrim(p_plugin.file_prefix,'/')||''')}';
    else
      -- try to find first IR/IG on the page
      for i in (SELECT nvl(static_id,'R'||to_char(region_id)) as affected_region_selector,
                         r.source_type
                  FROM apex_application_page_regions r
                  where r.page_id = nv('APP_PAGE_ID')
                    and r.application_id =nv('APP_ID')
                    and r.source_type  in ('Interactive Report','Interactive Grid')
                    and r.workspace = v_workspace
                    and rownum < 2
                 )
        loop
          if i.source_type = 'Interactive Report' then
            v_result.javascript_function := 'function(){excel_gpv.getExcel('''||i.affected_region_selector||''','''||v_plugin_id||''')}';
          else
            v_result.javascript_function := 'function(){excel_ig_gpv.downloadXLSXfromIG('''||i.affected_region_selector||''','''||v_plugin_id||''','''||get_ig_file_name(v_affected_region_IG_selector)||''',''/'||ltrim(p_plugin.file_prefix,'/')||''')}';
          end if; 
          v_found := true;
        end loop;        
        if not v_found then
          v_result.javascript_function := 'function(){console.log("No Affected Region Found!");}';
        end if;  
    end if;
    
    v_result.ajax_identifier := v_plugin_id;

    return v_result;
  end render;
  
  ------------------------------------------------------------------------------
  --used for export in IG
  procedure print_column_properties_json(p_application_id in number,
                                         p_page_id        in number
                                        )
  is
    l_columns_cursor    SYS_REFCURSOR;
    l_highlihts_cursor  SYS_REFCURSOR;
    v_decimal_separator char(1 char);
    v_lang_code         char(2 char);
  begin
    open l_columns_cursor for 
    select column_id,
           case 
            when  data_type in ('DATE','TIMESTAMP_TZ','TIMESTAMP_LTZ','TIMESTAMP') then 'DATE'
            else data_type
           end data_type,
           name,
           case 
            when  data_type in ('DATE','TIMESTAMP_TZ','TIMESTAMP_LTZ','TIMESTAMP') then
                  ir_to_xlsx.convert_date_format_js(p_datatype => data_type, p_format => format_mask)
            else ''
           end date_format_mask_js,
           case 
            when  data_type in ('DATE','TIMESTAMP_TZ','TIMESTAMP_LTZ','TIMESTAMP') then
                  ir_to_xlsx.convert_date_format(p_format => nvl(format_mask,'DD.MM.YYYY'))
            else ''
           end date_format_mask_excel,
           value_alignment,
           heading_alignment
    from APEX_APPL_PAGE_IG_COLUMNS
    where application_id = p_application_id 
      and page_id = p_page_id
    order by display_sequence;
    
    open l_highlihts_cursor for 
    select highlight_id,
           background_color,
           text_color
    from apex_appl_page_ig_rpt_highlts
    where application_id = p_application_id 
      and page_id = p_page_id;

    select substr(value,1,1)  as decimal_seperator
    into v_decimal_separator
    from nls_session_parameters
    where parameter = 'NLS_NUMERIC_CHARACTERS';
    
    -- always use 'AMERICA' as second parameter because 
    -- really i need only lang code (first parameter) and not the country
    select regexp_substr(UTL_I18N.MAP_LOCALE_TO_ISO  (value, 'AMERICA'),'[^_]+')
    into v_lang_code
    from nls_session_parameters
    where parameter = 'NLS_LANGUAGE';

    APEX_JSON.initialize_clob_output;
    APEX_JSON.open_object;
    APEX_JSON.write('column_properties', l_columns_cursor);
    APEX_JSON.write('highlights', l_highlihts_cursor);
    APEX_JSON.write('decimal_separator', v_decimal_separator);
    APEX_JSON.write('lang_code', v_lang_code);
    APEX_JSON.close_object;
    sys.htp.p(APEX_JSON.get_clob_output);
    APEX_JSON.free_output;
    
    if l_columns_cursor%ISOPEN THEN
       close l_columns_cursor;
    end if;
    if l_highlihts_cursor%ISOPEN THEN
       close l_highlihts_cursor;
    end if;    
  end print_column_properties_json;
  ------------------------------------------------------------------------------
  
  function ajax (p_dynamic_action in apex_plugin.t_dynamic_action,
                 p_plugin         in apex_plugin.t_plugin )
  return apex_plugin.t_dynamic_action_ajax_result
  is
    p_download_type      varchar2(1);
    p_custom_width       t_large_varchar2;
	p_autofilter         char;
    p_export_links       char;
    v_maximum_rows       number;
    v_dummy              apex_plugin.t_dynamic_action_ajax_result;
    v_affected_region_id apex_application_page_da_acts.affected_region_id%type;
  begin
      v_plugin_running := true;
      --to get properties needed for export in IG
      if apex_application.g_x01 = 'G' then 
        print_column_properties_json(p_application_id => apex_application.g_x02,
                                     p_page_id        => apex_application.g_x03);
        return v_dummy;
      end if;  
  
      p_download_type := nvl(p_dynamic_action.attribute_02,'E');
      p_autofilter := nvl(p_dynamic_action.attribute_04,'Y');
      p_export_links := nvl(p_dynamic_action.attribute_05,'N');
      p_custom_width := p_dynamic_action.attribute_06;
      v_affected_region_id := get_affected_region_id(p_dynamic_action_id => p_dynamic_action.ID
                                                    ,p_html_region_id    => apex_application.g_x03);
      
      v_maximum_rows := nvl(nvl(APEX_PLUGIN_UTIL.GET_PLSQL_EXPRESSION_RESULT(nvl(p_dynamic_action.attribute_01,' null ')),
                                IR_TO_XLSX.get_max_rows (p_app_id    => apex_application.g_x01,
                                                         p_page_id   => apex_application.g_x02,
                                                         p_region_id => v_affected_region_id)
                                ),1000);                                               
      if p_download_type = 'E' then 
        ir_to_xlsx.download_excel(p_app_id        => apex_application.g_x01,
                                  p_page_id      => apex_application.g_x02,
                                  p_region_id    => v_affected_region_id,
                                  p_col_length   => apex_application.g_x04,
                                  p_max_rows     => v_maximum_rows,
                                  p_autofilter   => p_autofilter,
                                  p_export_links => p_export_links,
                                  p_custom_width => p_custom_width
                                  );
      elsif p_download_type = 'T' then 
        ir_to_xlsx.download_debug(p_app_id        => apex_application.g_x01,
                                  p_page_id      => apex_application.g_x02,
                                  p_region_id    => v_affected_region_id,
                                  p_col_length   => apex_application.g_x04,
                                  p_max_rows     => v_maximum_rows,
                                  p_autofilter   => p_autofilter,
                                  p_export_links => p_export_links,
                                  p_custom_width => p_custom_width
                                  );
      elsif p_download_type = 'M' then -- use Moritz Klein engine https://github.com/commi235
        apexir_xlsx_pkg.download( p_ir_region_id => v_affected_region_id,
                                  p_app_id => apex_application.g_x01,
                                  p_ir_page_id => apex_application.g_x02
                                 );      
      else
        raise_application_error(-20001,'Unknown download_type: '||p_download_type);
      end if;
     return v_dummy;
  exception
    when others then
      raise_application_error(-20001,SQLERRM||chr(10)||dbms_utility.format_error_backtrace);      
  end ajax;
  
end IR_TO_MSEXCEL;
/

