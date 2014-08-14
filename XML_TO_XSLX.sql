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
