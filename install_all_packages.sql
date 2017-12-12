/**********************************************
**
** Author: Pavel Glebov
** Date: 10-2017
** Version: 2.21
**
** This all in one install script contains headrs and bodies of 5 packages
**
** IR_TO_XML.sql 
** AS_ZIP.sql  
** IR_TO_XSLX.sql
** IR_TO_MSEXCEL.sql
** APEXIR_XLSX_PKG.sql (mock)
**
**********************************************/

set define off;


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
CREATE OR REPLACE PACKAGE  "AS_ZIP" 
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


CREATE OR REPLACE PACKAGE BODY  "AS_ZIP" 
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

CREATE OR REPLACE PACKAGE BODY IR_TO_XLSX
as    

  STRING_HEIGHT      constant number default 14.4; 
  WIDTH_COEFFICIENT CONSTANT NUMBER := 6;
  subtype t_color is varchar2(7);
  subtype t_style_string  is varchar2(300);
  subtype t_format_mask  is varchar2(100);
  subtype t_font  is varchar2(50);
  subtype t_large_varchar2  is varchar2(32767);
  BACK_COLOR  constant  t_color default '#C6E0B4';
  
  type t_styles is table of binary_integer index by t_style_string;
  a_styles t_styles;
  type  t_color_list is table of binary_integer index by t_color;
  type  t_font_list is table of binary_integer index by t_font;
  a_font       t_font_list;
  a_back_color t_color_list;
  v_fonts_xml  t_large_varchar2;
  v_back_xml   t_large_varchar2;
  v_styles_xml clob;
  type  t_format_mask_list is table of binary_integer index by t_format_mask;
  a_format_mask_list t_format_mask_list;
  v_format_mask_xml clob;
  
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
  
  t_style_template clob default '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
    #FORMAT_MASKS#
    #FONTS#
    #FILLS#
    <borders count="2">
      <border>
        <left/>
        <right/>
        <top/>
        <bottom/>
        <diagonal/>
      </border>
      <border>
        <left/>  
        <right/> 
          <top style="thin">
        <color indexed="64"/>
        </top>
        <bottom/>
        <diagonal/>
      </border>
    </borders>
    <cellStyleXfs count="1">
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
    </cellStyleXfs>
    #STYLES#
    <cellStyles count="1">
      <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
    <extLst>
      <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
        <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
      </ext>
    </extLst>
  </styleSheet>';
  
  DEFAULT_FONT constant varchar2(200) := '
   <font>
      <sz val="11" />
      <color theme="1" />
      <name val="Calibri" />
      <family val="2" />
      <scheme val="minor" />
    </font>';
  BOLD_FONT constant varchar2(200) := '
   <font>
      <b />
      <sz val="11" />
      <color theme="1" />
      <name val="Calibri" />
      <family val="2" />
      <scheme val="minor" />
    </font>';
  FONTS_CNT constant  binary_integer := 2;   
  
  
  DEFAULT_FILL constant varchar2(200) :=  '
    <fill>
      <patternFill patternType="none" />
    </fill>
    <fill>
      <patternFill patternType="gray125" />
    </fill>'; 
  DEFAULT_FILL_CNT constant  binary_integer := 2; 
  
  DEFAULT_STYLE constant varchar2(200) := '
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
  AGGREGATE_STYLE constant varchar2(250) := '
      <xf numFmtId="#FMTID#" borderId="1" fillId="0" fontId="1" xfId="0" applyAlignment="1" applyFont="1" applyBorder="1">
         <alignment wrapText="1" horizontal="right"  vertical="top"/>
     </xf>';
  
     
  HEADER_STYLE constant varchar2(200) := '     
      <xf numFmtId="0" borderId="0" #FILL# fontId="1" xfId="0" applyAlignment="1" applyFont="1" >
         <alignment wrapText="0" horizontal="#ALIGN#"/>
     </xf>';  
  
  
  DEFAULT_STYLES_CNT constant  binary_integer := 1;     
  FORMAT_MASK_START_WITH  constant  binary_integer := 164;  
  
  FORMAT_MASK constant varchar2(100) := '
      <numFmt numFmtId="'||FORMAT_MASK_START_WITH||'" formatCode="dd.mm.yyyy"/>';
  
  FORMAT_MASK_CNT constant binary_integer default 1; 
  
/*
** Minor bugfixes by J.P.Lourens  9-Oct-2016
*/  
  subtype largevarchar2 is varchar2(32767); 
  subtype columntype is varchar2(15); 
  subtype formatmask is varchar2(100);  
  format_error EXCEPTION;
  PRAGMA EXCEPTION_INIT(format_error, -01830);
  date_format_error EXCEPTION;
  PRAGMA EXCEPTION_INIT(date_format_error, -01821);
  conversion_error exception;
  PRAGMA EXCEPTION_INIT(conversion_error,-06502);
  
  type t_date_table is table of date index by binary_integer;  
  type t_number_table is table of number index by binary_integer;  
  
  cursor cur_highlight(p_report_id in APEX_APPLICATION_PAGE_IR_RPT.REPORT_ID%TYPE,
                       p_delimetered_column_list in varchar2) 
  IS
  select rez.* ,
       rownum COND_NUMBER,
       'HIGHLIGHT_'||rownum COND_NAME
  from (
  select report_id,
         case when condition_operator in ('not in', 'in') then
                 IR_TO_XLSX.get_highlight_in_cond_sql(CONDITION_EXPRESSION,CONDITION_SQL,CONDITION_COLUMN_NAME)
              else 
                 replace(replace(replace(replace(condition_sql,'#APXWS_EXPR#',''''||CONDITION_EXPRESSION||''''),'#APXWS_EXPR2#',''''||CONDITION_EXPRESSION2||''''),'#APXWS_HL_ID#','1'),'#APXWS_CC_EXPR#','"'||CONDITION_COLUMN_NAME||'"') 
             end condition_sql,
         CONDITION_COLUMN_NAME,
         CONDITION_ENABLED,
         HIGHLIGHT_ROW_COLOR,
         HIGHLIGHT_ROW_FONT_COLOR,
         HIGHLIGHT_CELL_COLOR,
         HIGHLIGHT_CELL_FONT_COLOR      
    from APEX_APPLICATION_PAGE_IR_COND
    where condition_type = 'Highlight'
      and report_id = p_report_id
      and instr(':'||p_delimetered_column_list||':',':'||CONDITION_COLUMN_NAME||':') > 0
      and condition_enabled = 'Yes'
      order by --rows highlights first 
             nvl2(HIGHLIGHT_ROW_COLOR,1,0) desc, 
             nvl2(HIGHLIGHT_ROW_FONT_COLOR,1,0) desc,
             HIGHLIGHT_SEQUENCE 
    ) rez;
  
  type t_col_names is table of apex_application_page_ir_col.report_label%type index by apex_application_page_ir_col.column_alias%type;
  type t_col_format_mask is table of APEX_APPLICATION_PAGE_IR_COMP.computation_format_mask%TYPE index by APEX_APPLICATION_PAGE_IR_COL.column_alias%TYPE;
  type t_header_alignment is table of APEX_APPLICATION_PAGE_IR_COL.heading_alignment%TYPE index by APEX_APPLICATION_PAGE_IR_COL.column_alias%TYPE;
  type t_column_alignment is table of apex_application_page_ir_col.column_alignment%type index by apex_application_page_ir_col.column_alias%type;
  type t_column_types is table of apex_application_page_ir_col.column_type%type index by binary_integer;
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
    skipped_columns           BINARY_INTEGER default 0, -- when scpecial coluns like apxws_row_pk is used
    start_with                BINARY_INTEGER default 0, -- position of first displayed column in query
    end_with                  BINARY_INTEGER default 0, -- position of last displayed column in query    
    agg_cols_cnt              BINARY_INTEGER default 0, 
    hidden_cols_cnt           BINARY_INTEGER default 0, 
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
     text            largevarchar2,
     datatype        VARCHAR2(50)
   );  
  l_report                   ir_report;   
  v_debug                    clob;  
  v_debug_buffer             largevarchar2; 
  
 /* XLSX FUNCTIONS */
  function convert_date_format(p_format in varchar2)
  return varchar2
  is
    v_str      varchar2(100);
  begin
    v_str := p_format;
    v_str := upper(v_str);
    --date
    v_str := replace(v_str,'DAY','DDDD');
    v_str := replace(v_str,'MONTH','MMMM');
    v_str := replace(v_str,'MON','MMM');
    v_str := replace(v_str,'R','Y');    
    v_str := replace(v_str,'FM','');
    v_str := replace(v_str,'PM',' AM/PM');
    --time
    v_str := replace(v_str,'MI','mm');
    v_str := replace(v_str,'SS','ss');
    
    v_str := replace(v_str,'HH24','hh');
    --v_str := regexp_replace(v_str,'(\W)','\\\1');
    v_str := regexp_replace(v_str,'HH12([^ ]+)','h\1 AM/PM');    
    v_str := replace(v_str,'AM\/PM',' AM/PM');
    
    
    return v_str;
  end convert_date_format;
  ------------------------------------------------------------------------------
  function convert_number_format(p_format in varchar2)
  return varchar2
  is
    v_str      varchar2(100);
  begin
    v_str := p_format;
    v_str := upper(v_str);
    --number
    v_str := replace(v_str,'9','#');
    v_str := replace(v_str,'D','.');
    v_str := replace(v_str,'G',',');
    v_str := replace(v_str,'FM','');
    -- plus/minus sign
    if instr(v_str,'S') > 0 then
      v_str := replace(v_str,'S','');
      v_str := '+'||v_str||';-'||v_str;
    end if;
    
    --currency
    v_str := replace(v_str,'L',convert('&quot;'||rtrim(to_char(0,'FML0'),'0')||'&quot;','UTF8'));    
    
    v_str := regexp_substr(v_str,'.G?[^G]+$');
    return v_str;
  end convert_number_format;  
  ------------------------------------------------------------------------------
  
  function add_font(p_font in t_color,p_bold in varchar2 default null)
  return binary_integer 
  is  
   v_index t_font;
  begin
    v_index := p_font||p_bold;
    
    if not a_font.exists(v_index) then
      a_font(v_index) := a_font.count + 1;
      v_fonts_xml := v_fonts_xml||
        '<font>'||chr(10)||
        '   <sz val="11" />'||chr(10)||
        '   <color rgb="'||ltrim(p_font,'#')||'" />'||chr(10)||
        '   <name val="Calibri" />'||chr(10)||
        '   <family val="2" />'||chr(10)||
        '   <scheme val="minor" />'||
        '</font>'||chr(10);
        
        return a_font.count + FONTS_CNT - 1; --start with 0
     else
       return a_font(v_index) + FONTS_CNT - 1;
     end if;
  end add_font;
  ------------------------------------------------------------------------------
  
  function  add_back_color(p_back in t_color)
  return binary_integer 
  is
  begin
    if not a_back_color.exists(p_back) then
      a_back_color(p_back) := a_back_color.count + 1;
      v_back_xml := v_back_xml||
        '<fill>'||chr(10)||
        '   <patternFill patternType="solid">'||chr(10)||
        '     <fgColor rgb="'||ltrim(p_back,'#')||'" />'||chr(10)||
        '     <bgColor indexed="64" />'||chr(10)||
        '   </patternFill>'||
        '</fill>'||chr(10);
        
        return  a_back_color.count + DEFAULT_FILL_CNT - 1; --start with 0
     else
       return a_back_color(p_back) + DEFAULT_FILL_CNT - 1;
     end if;
  end add_back_color;
  ------------------------------------------------------------------------------
  function  add_format_mask(p_mask in varchar2)
  return binary_integer 
  is
  begin
    if not a_format_mask_list.exists(p_mask) then
      a_format_mask_list(p_mask) := a_format_mask_list.count + 1;
      v_format_mask_xml := v_format_mask_xml||
        '<numFmt numFmtId="'||(FORMAT_MASK_START_WITH + a_format_mask_list.count)||'" formatCode="'||p_mask||'"/>'||chr(10);
        return  a_format_mask_list.count + FORMAT_MASK_CNT + FORMAT_MASK_START_WITH - 1; 
     else
       return a_format_mask_list(p_mask) + FORMAT_MASK_CNT + FORMAT_MASK_START_WITH - 1;
     end if;
  end add_format_mask;  
  ------------------------------------------------------------------------------
  function get_font_colors_xml 
  return clob
  is
  begin
    return to_clob('<fonts count="'||(a_font.count + FONTS_CNT)||'" x14ac:knownFonts="1">'||
                   DEFAULT_FONT||chr(10)||
                   BOLD_FONT||chr(10)||
                   v_fonts_xml||chr(10)||
                   '</fonts>'||chr(10));
  end get_font_colors_xml;  
  ------------------------------------------------------------------------------
  function get_back_colors_xml 
  return clob
  is
  begin
    return to_clob('<fills count="'||(a_back_color.count + DEFAULT_FILL_CNT)||'">'||
                   DEFAULT_FILL||
                   v_back_xml||chr(10)||
                   '</fills>'||chr(10));
  end get_back_colors_xml;  
  ------------------------------------------------------------------------------  
  function get_format_mask_xml 
  return clob
  is
  begin
    return to_clob('<numFmts count="'||(a_format_mask_list.count + FORMAT_MASK_CNT)||'">'||
                   FORMAT_MASK||
                   v_format_mask_xml||chr(10)||
                   '</numFmts>'||chr(10));
  end get_format_mask_xml;  
  ------------------------------------------------------------------------------  
  function get_cellXfs_xml
  return clob
  is
  begin
    return to_clob('<cellXfs count="'||(a_styles.count + DEFAULT_STYLES_CNT)||'">'||
                   DEFAULT_STYLE||chr(10)||
                   v_styles_xml||chr(10)||
                   '</cellXfs>'||chr(10));
  end get_cellXfs_xml;
  ------------------------------------------------------------------------------  
  function get_num_fmt_id(p_data_type   in varchar2,
                          p_format_mask in varchar2)
  return binary_integer
  is
  begin
      if p_data_type = 'NUMBER' then 
       if p_format_mask is null then
          return 0;
        else
          return add_format_mask(convert_number_format(p_format_mask)); 
        end if;
      elsif p_data_type = 'DATE' then 
        if p_format_mask is null then
          return FORMAT_MASK_START_WITH;  -- default date format
        else
          return add_format_mask(convert_date_format(p_format_mask)); 
        end if;
      else
        return 2;  -- default  string format
      end if;  
  
  end get_num_fmt_id;  
  ------------------------------------------------------------------------------  
  -- get style_id for existent styles or 
  -- add new style and return style_id
  function get_style_id(p_font        in t_color,
                        p_back        in t_color,
                        p_data_type   in varchar2,
                        p_format_mask in varchar2,
                        p_align       in varchar2)
  return binary_integer
  is
    v_style           t_style_string;
    v_font_color_id   binary_integer;
    v_back_color_id   binary_integer;
    v_style_xml       t_large_varchar2;
  begin
    v_style := nvl(p_font,'      ')||nvl(p_back,'       ')||p_data_type||p_format_mask||p_align;
   
    --   
    if a_styles.exists(v_style) then
      return a_styles(v_style)   - 1 + DEFAULT_STYLES_CNT;
    else
      a_styles(v_style) := a_styles.count + 1;
      
      v_style_xml := '<xf borderId="0" xfId="0" ';
      v_style_xml := v_style_xml||replace(' numFmtId="#FMTID#" ','#FMTID#',get_num_fmt_id(p_data_type,p_format_mask));
      
      if p_font is not null then
        v_font_color_id := add_font(p_font);
        v_style_xml := v_style_xml||' fontId="'||v_font_color_id||'"  applyFont="1" ';
      else
        v_style_xml := v_style_xml||' fontId="0" '; --default font
      end if;
      
      if p_back is not null then
        v_back_color_id := add_back_color(p_back);
        v_style_xml := v_style_xml||' fillId="'||v_back_color_id||'"  applyFill="1" ';
      else
        v_style_xml := v_style_xml||' fillId="0" '; --default fill 
      end if;
      
      v_style_xml := v_style_xml||'>'||chr(10);
      v_style_xml := v_style_xml||'<alignment wrapText="1"';
      if p_align is not null then
        v_style_xml := v_style_xml||' horizontal="'||lower(p_align)||'" ';
      end if;
      v_style_xml := v_style_xml||'/>'||chr(10);
      
      v_style_xml := v_style_xml||'</xf>'||chr(10);      
      v_styles_xml := v_styles_xml||to_clob(v_style_xml);     
       
      return a_styles.count  - 1 + DEFAULT_STYLES_CNT;
    end if;  
  
  end get_style_id;
  ------------------------------------------------------------------------------  
  -- get style_id for existent styles or 
  -- add new style and return style_id
  function get_aggregate_style_id(p_font        in t_color,                               
                                  p_back        in t_color,
                                  p_data_type   in varchar2,
                                  p_format_mask in varchar2)
  return binary_integer
  is
    v_style           t_style_string;
    v_style_xml       t_large_varchar2;
    v_num_fmt_id        binary_integer;
  begin
    v_style := nvl(p_font,'      ')||nvl(p_back,'       ')||p_data_type||p_format_mask||'AGGREGATE';
    --   
    if a_styles.exists(v_style) then
      return a_styles(v_style)  - 1 + DEFAULT_STYLES_CNT;
    else
      a_styles(v_style) := a_styles.count + 1;
      v_style_xml  := replace(AGGREGATE_STYLE,'#FMTID#',get_num_fmt_id(p_data_type,p_format_mask)) ||chr(10);      
      v_styles_xml := v_styles_xml||v_style_xml;      
      return a_styles.count  - 1 + DEFAULT_STYLES_CNT;
    end if;  
  
  end get_aggregate_style_id;  
  ------------------------------------------------------------------------------
  function get_header_style_id(p_back        in t_color default BACK_COLOR,
                               p_align       in varchar2,
                               p_border      in boolean default false)
  return binary_integer
  is
    v_style             t_style_string;
    v_style_xml         t_large_varchar2;
    v_num_fmt_id        binary_integer;
    v_back_color_id     binary_integer;
  begin
    v_style := '      '||nvl(p_back,'       ')||'CHARHEADER'||p_align;
    --   
    if a_styles.exists(v_style) then
      return a_styles(v_style)  - 1 + DEFAULT_STYLES_CNT;
    else
      a_styles(v_style) := a_styles.count + 1;
  
      if p_back is not null then
        v_back_color_id := add_back_color(p_back);
        v_style_xml := replace(HEADER_STYLE,'#FILL#',' fillId="'||v_back_color_id||'"  applyFill="1" ' );
      else
        v_style_xml := replace(HEADER_STYLE,'#FILL#',' fillId="0" '); --default fill 
      end if;      
      v_style_xml  := replace(v_style_xml,'#ALIGN#',lower(p_align))||chr(10);
      v_styles_xml := v_styles_xml||v_style_xml;      
      return a_styles.count  - 1 + DEFAULT_STYLES_CNT;
    end if;  
  
  end get_header_style_id;    
  ------------------------------------------------------------------------------
  
  function  get_styles_xml
  return clob
  is
    v_template clob;
  begin
     v_template := replace(t_style_template,'#FORMAT_MASKS#',get_format_mask_xml);     
     v_template := replace(v_template,'#FONTS#',get_font_colors_xml);
     v_template := replace(v_template,'#FILLS#',get_back_colors_xml);
     v_template := replace(v_template,'#STYLES#',get_cellXfs_xml);
     
     return v_template;
  end get_styles_xml;
  ------------------------------------------------------------------------------
  /**
  * http://mk-commi.blogspot.co.at/2014/11/concatenating-varchar2-values-into-clob.html  
  
  * Procedure concatenates a VARCHAR2 to a CLOB.
  * It uses another VARCHAR2 as a buffer until it reaches 32767 characters.
  * Then it flushes the current buffer to the CLOB and resets the buffer using
  * the actual VARCHAR2 to add.
  * Your final call needs to be done setting p_eof to TRUE in order to
  * flush everything to the CLOB.
  *
  * @param p_clob        The CLOB buffer.
  * @param p_vc_buffer   The intermediate VARCHAR2 buffer. (must be VARCHAR2(32767))
  * @param p_vc_addition The VARCHAR2 value you want to append.
  * @param p_eof         Indicates if complete buffer should be flushed to CLOB.
  */
  PROCEDURE add( p_clob IN OUT NOCOPY CLOB
               , p_vc_buffer IN OUT NOCOPY VARCHAR2
               , p_vc_addition IN VARCHAR2
               , p_eof IN BOOLEAN DEFAULT FALSE
               )
  AS
  BEGIN
     
    -- Standard Flow
    IF NVL(LENGTHB(p_vc_buffer), 0) + NVL(LENGTHB(p_vc_addition), 0) < (32767/2) THEN
      -- Danke f?r Frank Menne wegen utf-8
      p_vc_buffer := p_vc_buffer || convert(p_vc_addition,'utf8');
    ELSE
      IF p_clob IS NULL THEN
        dbms_lob.createtemporary(p_clob, TRUE);
      END IF;
      dbms_lob.writeappend(p_clob, length(p_vc_buffer), p_vc_buffer);
      p_vc_buffer := convert(p_vc_addition,'utf8');
    END IF;
     
    -- Full Flush requested
    IF p_eof THEN
      IF p_clob IS NULL THEN
         p_clob := p_vc_buffer;
      ELSE
        dbms_lob.writeappend(p_clob, length(p_vc_buffer), p_vc_buffer);
      END IF;
      p_vc_buffer := NULL;
    END IF;
   
  END add;
  ------------------------------------------------------------------------------
  FUNCTION get_cell_name(p_col IN binary_integer,
                         p_row IN binary_integer)
  RETURN varchar2
  IS   
  BEGIN
  /*
   Author: Moritz Klein (https://github.com/commi235)
   https://github.com/commi235/xlsx_builder/blob/master/xlsx_builder_pkg.pkb
  */
      RETURN CASE
                WHEN p_col > 702
                THEN
                      CHR (64 + TRUNC ( (p_col - 27) / 676))
                   || CHR (65 + MOD (TRUNC ( (p_col - 1) / 26) - 1, 26))
                   || CHR (65 + MOD (p_col - 1, 26))
                WHEN p_col > 26
                THEN
                   CHR (64 + TRUNC ( (p_col - 1) / 26)) || CHR (65 + MOD (p_col - 1, 26))
                ELSE
                   CHR (64 + p_col)
             END||p_row;
  end get_cell_name;
 -----------------------------------------------------------
  procedure log(p_message in varchar2,p_eof IN BOOLEAN DEFAULT FALSE)
  is
  begin
  /* logigging  ffrffe */
    add(v_debug,v_debug_buffer,p_message||chr(10),p_eof);
    apex_debug_message.log_message(p_message => substr(p_message,1,32767),
                                   p_enabled => false,
                                   p_level   => 4);
  end log; 
  ------------------------------------------------------------------------------
  
  function get_column_names(p_column_alias in apex_application_page_ir_col.column_alias%type)
  return APEX_APPLICATION_PAGE_IR_COL.report_label%TYPE
  is
  begin
    -- https://github.com/glebovpavel/IR_to_MSExcel/issues/9
    -- Thanks HeavyS 
    return  apex_plugin_util.replace_substitutions(p_value => l_report.column_names(p_column_alias));
  exception
    when others then
       raise_application_error(-20001,'get_column_names: p_column_alias='||p_column_alias||' '||SQLERRM);
  end get_column_names;
  ------------------------------------------------------------------------------
  function get_col_format_mask(p_column_alias in apex_application_page_ir_col.column_alias%type)
  return formatmask
  is
  begin
    return replace(l_report.col_format_mask(p_column_alias),'"','');
  exception
    when others then
       raise_application_error(-20001,'get_col_format_mask: p_column_alias='||p_column_alias||' '||SQLERRM);
  end get_col_format_mask;
  ------------------------------------------------------------------------------
  procedure set_col_format_mask(p_column_alias in apex_application_page_ir_col.column_alias%type,
                                p_format_mask  in formatmask)
  is
  begin
      l_report.col_format_mask(p_column_alias) := p_format_mask;
  exception
    when others then
       raise_application_error(-20001,'set_col_format_mask: p_column_alias='||p_column_alias||' '||SQLERRM);
  end set_col_format_mask;
  ------------------------------------------------------------------------------
  function get_header_alignment(p_column_alias in apex_application_page_ir_col.column_alias%type)
  return APEX_APPLICATION_PAGE_IR_COL.heading_alignment%TYPE
  is
  begin
    return l_report.header_alignment(p_column_alias);
  exception
    when others then
       raise_application_error(-20001,'get_header_alignment: p_column_alias='||p_column_alias||' '||SQLERRM);
  end get_header_alignment;
  ------------------------------------------------------------------------------
  function get_column_alignment(p_column_alias in apex_application_page_ir_col.column_alias%type)
  return apex_application_page_ir_col.column_alignment%type
  is
  begin
    return l_report.column_alignment(p_column_alias);
  exception
    when others then
       raise_application_error(-20001,'get_column_alignment: p_column_alias='||p_column_alias||' '||SQLERRM);
  end get_column_alignment;
  ------------------------------------------------------------------------------
  function get_column_types(p_num in binary_integer)
  return apex_application_page_ir_col.column_type%type
  is
  begin
    return l_report.column_types(p_num);
  exception
    when others then
       raise_application_error(-20001,'get_column_names: p_num='||p_num||' '||SQLERRM);
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
                           p_id          in binary_integer)
  return largevarchar2
  is
  begin
    return p_current_row(p_id);
  exception
    when others then
       raise_application_error(-20001,'get_current_row: string: p_id='||p_id||' '||SQLERRM);
  end get_current_row; 
  ------------------------------------------------------------------------------
  function get_current_row(p_current_row in t_date_table,
                           p_id          in binary_integer)
  return date
  is
  begin
    if p_current_row.exists(p_id) then
      return p_current_row(p_id);
    else
      return null;
    end if;
  exception
    when others then
       raise_application_error(-20001,'get_current_row:date: p_id='||p_id||' '||SQLERRM);
  end get_current_row;   
  ------------------------------------------------------------------------------
  function get_current_row(p_current_row in t_number_table,
                           p_id          in binary_integer)
  return number
  is
  begin
    if p_current_row.exists(p_id) then
      return p_current_row(p_id);
    else
      return null;
    end if;
  exception
    when others then
       raise_application_error(-20001,'get_current_row:number: p_id='||p_id||' '||SQLERRM);
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
    v_tmp largevarchar2;
  begin
    -- p_str can be encoded html-string 
    -- wee need first convert to text
    v_tmp := REGEXP_REPLACE(p_str,'<(BR)\s*/*>',chr(13)||chr(10),1,0,'i');
    v_tmp := REGEXP_REPLACE(v_tmp,'<[^<>]+>',' ',1,0,'i');
    -- https://community.oracle.com/message/14074217#14074217
    v_tmp := regexp_replace(v_tmp, '[^[:print:]'||chr(13)||chr(10)||chr(9)||']', ' ');
    v_tmp := UTL_I18N.UNESCAPE_REFERENCE(v_tmp); 
    -- and finally encode them
    v_tmp := substr(v_tmp,1,2000);    
    v_tmp := UTL_I18N.ESCAPE_REFERENCE(v_tmp,'UTF8');
    return v_tmp;
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
  function get_query_column_list
  return APEX_APPLICATION_GLOBAL.VC_ARR2
  is
   v_cur         INTEGER; 
   v_colls_count BINARY_INTEGER; 
   v_columns     APEX_APPLICATION_GLOBAL.VC_ARR2;
   v_desc_tab    DBMS_SQL.DESC_TAB2;
   v_sql         largevarchar2;   
  begin
    v_cur := dbms_sql.open_cursor(2);     
    v_sql := apex_plugin_util.replace_substitutions(p_value => l_report.report.sql_query,p_escape => false);
    log(v_sql);
    dbms_sql.parse(v_cur,v_sql,dbms_sql.native);     
    dbms_sql.describe_columns2(v_cur,v_colls_count,v_desc_tab);    
    for i in 1..v_colls_count loop
         if upper(v_desc_tab(i).col_name) != 'APXWS_ROW_PK' then --skip internal primary key if need
           v_columns(v_columns.count + 1) := v_desc_tab(i).col_name;
           log('Query column = '||v_desc_tab(i).col_name);
         end if;
    end loop;                 
   dbms_sql.close_cursor(v_cur);   
   return v_columns;
  exception
    when others then
      if dbms_sql.is_open(v_cur) then
        dbms_sql.close_cursor(v_cur);
      end if;  
      raise_application_error(-20001,'get_query_column_list: '||SQLERRM);
  end get_query_column_list;  
  ------------------------------------------------------------------------------
  function get_cols_as_table(p_delimetered_column_list     in varchar2,
                             p_displayed_nonbreak_columns  in apex_application_global.vc_arr2)
  return apex_application_global.vc_arr2
  is
  begin
    return intersect_arrays(APEX_UTIL.STRING_TO_TABLE(rr(p_delimetered_column_list)),p_displayed_nonbreak_columns);
  end get_cols_as_table;
  
  ------------------------------------------------------------------------------
  function get_hidden_columns_cnt(p_app_id       in number,
                                  p_page_id      in number,
                                  p_region_id    in number,
                                  p_report_id    in number)
  return number
  -- J.P.Lourens 9-Oct-16 added p_region_id as input variable, and added v_get_query_column_list
  is 
   v_hidden_columns  number default 0;
   v_hidden_computation_columns  number default 0;
   v_get_query_column_list varchar2(32676);
  begin
  
      v_get_query_column_list := apex_util.table_to_string(get_query_column_list);
      
      select count(*)
      into v_hidden_columns
      from APEX_APPLICATION_PAGE_IR_COL
     where application_id = p_app_id
       AND page_id = p_page_id
       -- J.P.Lourens 9-Oct-16 added p_region_id to ensure correct results when having multiple IR on a page
       and region_id = p_region_id
       and (display_text_as = 'HIDDEN'
       -- J.P.Lourens 9-Oct-2016 modified get_hidden_columns_cnt to INCLUDE columns which are
       --                        - selected in the IR query, and thus included in v_get_query_column_list
       --                        - not included in the report, thus missing from l_report.ir_data.report_columns
          or instr(':'||l_report.ir_data.report_columns||':',':'||column_alias||':') = 0)
       and instr(':'||v_get_query_column_list||':',':'||column_alias||':') > 0;
       
       select count(*)
       into v_hidden_computation_columns
       from apex_application_page_ir_comp
       where application_id = p_app_id
         and page_id = p_page_id       
         and report_id = p_report_id         
         and instr(':'||l_report.ir_data.report_columns||':',':'||computation_column_alias||':') = 0;       
      
      return v_hidden_columns + v_hidden_computation_columns;
  exception
    when no_data_found then
      return 0;
  end get_hidden_columns_cnt;
  ------------------------------------------------------------------------------ 
  
  procedure init_t_report(p_app_id       IN NUMBER,
                          p_page_id      IN NUMBER,
                          p_region_id    IN NUMBER)
  is
    l_report_id     number;
    v_query_targets apex_application_global.vc_arr2;
    l_new_report    ir_report; 
  begin
    l_report := l_new_report;
    --get base report id    
    log('l_region_id='||p_region_id);
    
    l_report_id := apex_ir.get_last_viewed_report_id (p_page_id   => p_page_id,
                                                      p_region_id => p_region_id);
    
    log('l_base_report_id='||l_report_id);    
    
    select r.* 
    into l_report.ir_data       
    from apex_application_page_ir_rpt r
    where application_id = p_app_id 
      and page_id = p_page_id
      and session_id = nv('APP_SESSION')
      and application_user = v('APP_USER')
      and base_report_id = l_report_id;
  
    log('l_report_id='||l_report_id);
    l_report_id := l_report.ir_data.report_id;                                                                 
      
      
    l_report.report := apex_ir.get_report (p_page_id        => p_page_id,
                                           p_region_id      => p_region_id
                                          );
    l_report.ir_data.report_columns := APEX_UTIL.TABLE_TO_STRING(get_cols_as_table(l_report.ir_data.report_columns,get_query_column_list));
    
    -- J.P.Lourens 9-Oct-16 added p_region_id as input variable
    l_report.hidden_cols_cnt := get_hidden_columns_cnt(p_app_id    => p_app_id,
                                                       p_page_id   => p_page_id,
                                                       p_region_id => p_region_id,
                                                       p_report_id => l_report_id);
    
    <<displayed_columns>>                                      
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
                 and region_id = p_region_id
                 and display_text_as != 'HIDDEN' --l_report.ir_data.report_columns can include HIDDEN columns
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
      --l_report.column_types(i.column_alias) := i.column_type;
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
    end loop displayed_columns;    
    
    -- calculate columns count with aggregation separately
    l_report.sum_columns_on_break := get_cols_as_table(l_report.ir_data.sum_columns_on_break,l_report.displayed_columns);  
    l_report.avg_columns_on_break := get_cols_as_table(l_report.ir_data.avg_columns_on_break,l_report.displayed_columns);  
    l_report.max_columns_on_break := get_cols_as_table(l_report.ir_data.max_columns_on_break,l_report.displayed_columns);  
    l_report.min_columns_on_break := get_cols_as_table(l_report.ir_data.min_columns_on_break,l_report.displayed_columns);  
    l_report.median_columns_on_break := get_cols_as_table(l_report.ir_data.median_columns_on_break,l_report.displayed_columns); 
    l_report.count_columns_on_break := get_cols_as_table(l_report.ir_data.count_columns_on_break,l_report.displayed_columns);  
    l_report.count_distnt_col_on_break := get_cols_as_table(l_report.ir_data.count_distnt_col_on_break,l_report.displayed_columns); 
      
    -- calculate total count of columns with aggregation
    l_report.agg_cols_cnt := l_report.sum_columns_on_break.count + 
                             l_report.avg_columns_on_break.count +
                             l_report.max_columns_on_break.count + 
                             l_report.min_columns_on_break.count +
                             l_report.median_columns_on_break.count +
                             l_report.count_columns_on_break.count +
                             l_report.count_distnt_col_on_break.count;
    
    log('l_report.report_columns='||rr(l_report.ir_data.report_columns));    
    log('l_report.break_on='||rr(l_report.ir_data.break_enabled_on));
    log('l_report.sum_columns_on_break='||rr(l_report.ir_data.sum_columns_on_break));
    log('l_report.avg_columns_on_break='||rr(l_report.ir_data.avg_columns_on_break));
    log('l_report.max_columns_on_break='||rr(l_report.ir_data.max_columns_on_break));
    LOG('l_report.min_columns_on_break='||rr(l_report.ir_data.min_columns_on_break));
    log('l_report.median_columns_on_break='||rr(l_report.ir_data.median_columns_on_break));
    log('l_report.count_columns_on_break='||rr(l_report.ir_data.count_columns_on_break));    
    log('l_report.count_distnt_col_on_break='||rr(l_report.ir_data.count_distnt_col_on_break));
    log('l_report.break_really_on='||APEX_UTIL.TABLE_TO_STRING(l_report.break_really_on));
    log('l_report.agg_cols_cnt='||l_report.agg_cols_cnt);
    log('l_report.hidden_cols_cnt='||l_report.hidden_cols_cnt);
    
    
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
        v_query_targets(v_query_targets.count + 1) := c.condition_sql||' as HLIGHTS_'||(v_query_targets.count + 1);
    end loop;    
        
    if v_query_targets.count  > 0 then
      -- uwr485kv is random name 
      l_report.report.sql_query := 'SELECT '||APEX_UTIL.TABLE_TO_STRING(v_query_targets,','||chr(10))||', uwr485kv.* from ('||l_report.report.sql_query||') uwr485kv';
    end if;
    l_report.report.sql_query := l_report.report.sql_query;
    log('l_report.report.sql_query='||chr(10)||l_report.report.sql_query||chr(10));
  exception
    when no_data_found then
      raise_application_error(-20001,'No Interactive Report found on Page='||p_page_id||' Application='||p_app_id||' Please make sure that the report was running at least once by this session.');
    /*
    when others then 
     log('Exception in init_t_report');
     log('Page='||p_page_id);
     log('Application='||p_app_id);
     raise;
    */ 
  end init_t_report;  
  ------------------------------------------------------------------------------
 
  function is_control_break(p_curr_row  IN APEX_APPLICATION_GLOBAL.VC_ARR2,
                            p_prev_row  IN APEX_APPLICATION_GLOBAL.VC_ARR2)
  return boolean
  is
    v_start_with      BINARY_INTEGER;
    v_end_with        BINARY_INTEGER;    
  begin
    if nvl(l_report.break_really_on.count,0) = 0  then
      return false; --no control break
    end if;
    v_start_with := 1 + 
                    l_report.skipped_columns + 
                    l_report.row_highlight.count + 
                    l_report.col_highlight.count;    
    v_end_with   := l_report.skipped_columns + 
                    l_report.row_highlight.count + 
                    l_report.col_highlight.count +
                    nvl(l_report.break_really_on.count,0);
    for i in v_start_with..v_end_with loop
      if p_curr_row(i) != p_prev_row(i) then
        return true;
      end if;
    end loop;
    return false;
  end is_control_break;
  ------------------------------------------------------------------------------
  function get_formatted_number(p_number         IN number,
                                p_format_string  IN varchar2,
                                p_nls            IN varchar2 default null)
  return varchar2
  is
    v_str varchar2(100);
  begin
    v_str := trim(to_char(p_number,p_format_string,p_nls));
    if instr(v_str,'#') > 0 and ltrim(v_str,'#') is null then --format fail
      raise invalid_number;
    else
      return v_str;
    end if;    
  end get_formatted_number;  
  ------------------------------------------------------------------------------
  FUNCTION get_cell(p_query_value IN varchar2,
                    p_format_mask IN varchar2,
                    p_date        IN date)
  RETURN t_cell_data
  IS
    v_data t_cell_data;
  BEGIN     
     v_data.value := get_formatted_number(p_date - to_date('01-03-1900','DD-MM-YYYY') + 61,'9999999999999990D00000000','NLS_NUMERIC_CHARACTERS = ''.,''');     
     v_data.datatype := 'DATE';
     
     -- https://github.com/glebovpavel/IR_to_MSExcel/issues/16
     -- thanks Valentine Nikitsky 
     if p_format_mask is not null then
       if upper(p_format_mask)='SINCE' then
            v_data.text := apex_util.get_since(p_date);
            v_data.value := null;
            v_data.datatype := 'STRING';
       else
          v_data.text := to_char(p_date,p_format_mask);  -- normally v_data.text used in XML only
       end if;
     else
       v_data.text := p_query_value;
     end if;
     
     return v_data;
  EXCEPTION
    WHEN invalid_number or format_error THEN 
      v_data.value := NULL;          
      v_data.datatype := 'STRING';
      v_data.text := p_query_value;
      return v_data;
  END get_cell;  
  ------------------------------------------------------------------------------
  FUNCTION get_cell(p_query_value IN varchar2,
                    p_format_mask IN varchar2,
                    p_number      IN number)
  RETURN t_cell_data
  IS
    v_data t_cell_data;
  BEGIN
   v_data.datatype := 'NUMBER';   
   v_data.value := get_formatted_number(p_number,'9999999999999990D00000000','NLS_NUMERIC_CHARACTERS = ''.,''');
   
   if p_format_mask is not null then
     v_data.text := get_formatted_number(p_query_value,p_format_mask);
   ELSE
     v_data.text := p_query_value;
   end if;
   
   return v_data;
  EXCEPTION
    WHEN invalid_number or conversion_error THEN 
      v_data.value := NULL;          
      v_data.datatype := 'STRING';
      v_data.text := p_query_value;
      return v_data;
  END get_cell;  
  ------------------------------------------------------------------------------
     
       
  function print_control_break_header_obj(p_current_row     in apex_application_global.vc_arr2) 
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
      v_cb_xml := v_cb_xml ||
                  get_column_names(l_report.break_really_on(i))||': '||
                  get_current_row(p_current_row,i + 
                                                l_report.skipped_columns + 
                                                l_report.row_highlight.count + 
                                                l_report.col_highlight.count
                                 )||',';
    end loop visible_columns;
    
    return  rtrim(v_cb_xml,',');
  end print_control_break_header_obj;
  ------------------------------------------------------------------------------
  function find_rel_position (p_curr_col_name    IN varchar2,
                              p_agg_rows         IN APEX_APPLICATION_GLOBAL.VC_ARR2)
  return BINARY_INTEGER
  is
    v_relative_position BINARY_INTEGER;
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
  function get_agg_text(p_curr_col_name          IN VARCHAR2,
                        p_agg_rows               IN APEX_APPLICATION_GLOBAL.VC_ARR2,
                        p_current_row            IN APEX_APPLICATION_GLOBAL.VC_ARR2,
                        p_agg_text               IN VARCHAR2,
                        p_position               IN BINARY_INTEGER, --start position in sql-query
                        p_col_number             IN BINARY_INTEGER, --column position when displayed
                        p_default_format_mask    IN VARCHAR2 DEFAULT NULL,
                        p_overwrite_format_mask   IN varchar2 default null) --should be used forcibly 
  return varchar2
  is
    v_tmp_pos       BINARY_INTEGER;  -- current column position in sql-query 
    v_format_mask   apex_application_page_ir_comp.computation_format_mask%type;
    v_agg_value     largevarchar2;
    v_row_value     largevarchar2;
    v_g_format_mask varchar2(100);  
    v_col_alias     varchar2(255);
  begin
      v_tmp_pos := find_rel_position (p_curr_col_name,p_agg_rows); 
      if v_tmp_pos is not null then
        v_col_alias := get_column_alias_sql(p_col_number);
        v_g_format_mask :=  get_col_format_mask(v_col_alias);   
        v_format_mask := nvl(v_g_format_mask,p_default_format_mask);
        v_format_mask := nvl(p_overwrite_format_mask,v_format_mask);
        v_row_value :=  get_current_row(p_current_row,p_position + l_report.hidden_cols_cnt + v_tmp_pos);
        v_agg_value := trim(to_char(v_row_value,v_format_mask));
        
        return  get_xmlval(p_agg_text||v_agg_value||' '||chr(10));
      else
        return  '';
      end if;    
  exception
     when others then
        log('!Exception in get_agg_text');
        log('p_col_number='||p_col_number);
        log('v_col_alias='||v_col_alias);
        log('v_g_format_mask='||v_g_format_mask);
        log('p_default_format_mask='||p_default_format_mask);
        log('v_tmp_pos='||v_tmp_pos);
        log('p_position='||p_position);
        -- J.P.Lourens 9-Oct-16 added to log     
        log('l_report.hidden_cols_cnt='||l_report.hidden_cols_cnt);    
        log('v_row_value='||v_row_value);
        log('v_format_mask='||v_format_mask);
        raise;
  end get_agg_text;
  ------------------------------------------------------------------------------
  
  function get_aggregate(p_current_row     IN APEX_APPLICATION_GLOBAL.VC_ARR2) 
  return APEX_APPLICATION_GLOBAL.VC_ARR2
  is
    v_aggregate_xml   largevarchar2;
    v_agg_obj  APEX_APPLICATION_GLOBAL.VC_ARR2;
    v_position        BINARY_INTEGER;    
    v_i NUMBER := 0;
  begin
    if l_report.agg_cols_cnt  = 0 then      
      return v_agg_obj;
    end if;    
 
    <<visible_columns>>
    for i in l_report.start_with..l_report.end_with loop
      v_aggregate_xml := '';
      v_i := v_i + 1;
      v_position := l_report.end_with; --aggregate are placed after displayed columns and computations      
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
                                       p_col_number    => i,
                                       p_overwrite_format_mask   => '999G999G999G999G990');
      v_position := v_position + l_report.count_columns_on_break.count;                                 
      v_aggregate_xml := v_aggregate_xml || get_agg_text(p_curr_col_name => get_column_alias_sql(i),
                                       p_agg_rows      => l_report.count_distnt_col_on_break,
                                       p_current_row   => p_current_row,
                                       p_agg_text      => 'Count distinct:',
                                       p_position      => v_position,
                                       p_col_number    => i,
                                       p_overwrite_format_mask   => '999G999G999G999G990');
      v_agg_obj(v_i) := v_aggregate_xml;
    end loop visible_columns;
    return  v_agg_obj;
  end get_aggregate;        
  ------------------------------------------------------------------------------
  function can_show_as_date(p_format_string in varchar2)
  return boolean
  is 
    v_dummy varchar2(50);
  begin
    v_dummy := to_char(sysdate,p_format_string);
    
    return true;
  exception
    when invalid_number or format_error or date_format_error or conversion_error then 
      return false;
  end can_show_as_date;  
  ------------------------------------------------------------------------------
  function get_current_format(p_data_type in binary_integer)
  return varchar2
  is
    v_format    formatmask;
    v_parameter varchar2(50);
  begin
    if p_data_type in (dbms_types.TYPECODE_TIMESTAMP_TZ,181) then
       v_parameter := 'NLS_TIMESTAMP_TZ_FORMAT';
    elsif p_data_type in (dbms_types.TYPECODE_TIMESTAMP_LTZ,231) then
       v_parameter := 'NLS_TIMESTAMP_TZ_FORMAT';
    elsif p_data_type in (dbms_types.TYPECODE_TIMESTAMP,180) then
       v_parameter := 'NLS_TIMESTAMP_FORMAT';      
    elsif p_data_type = dbms_types.TYPECODE_DATE then
       v_parameter := 'NLS_DATE_FORMAT';      
    else 
       return 'dd.mm.yyyy';
    end if;     
    
    SELECT value 
    into v_format
    FROM V$NLS_Parameters 
    WHERE parameter = v_parameter;    
    
    return v_format;
  end get_current_format;
  ------------------------------------------------------------------------------
  -- excel has only DATE-format mask (no timezones etc)
  -- if format mask can be shown in excel - show column as date- type
  -- else - as string 
  procedure prepare_col_format_mask(p_col_number in binary_integer,
                                    p_data_type  in binary_integer )
  is
    v_format_mask          formatmask;
    v_default_format_mask  formatmask;
    v_final_format_mask    formatmask;
    v_col_alias            varchar2(255);    
  begin
     v_default_format_mask := get_current_format(p_data_type);
     log('v_default_format_mask='||v_default_format_mask);
     begin
       v_col_alias   := get_column_alias_sql(p_col_number);
       v_format_mask := get_col_format_mask(v_col_alias);
       log('v_col_alias='||v_col_alias||' v_format_mask='||v_format_mask);
     exception
       when no_data_found then 
          v_format_mask := '';
     end;  
     
     v_final_format_mask := nvl(v_format_mask,v_default_format_mask);
     if can_show_as_date(v_final_format_mask) then
        log('Can show as date');
        if v_col_alias is not null then 
           set_col_format_mask(p_column_alias => v_col_alias,
                               p_format_mask  => v_final_format_mask);
        end if;                       
        l_report.column_types(p_col_number) := 'DATE';
     else
        log('Can not show as date');
        l_report.column_types(p_col_number) := 'STRING';
     end if;
  end prepare_col_format_mask;
  ------------------------------------------------------------------------------
    
  procedure generate_from_report(v_clob in out nocopy clob,
                            v_strings_clob in out nocopy clob,
                            p_width_str in varchar2,
                            p_coefficient  in number,
                            p_max_rows in integer,
                            p_autofilter in char) 
  is
   v_cur                INTEGER; 
   v_result             INTEGER;
   v_colls_count        BINARY_INTEGER;
   v_row                APEX_APPLICATION_GLOBAL.VC_ARR2;
   v_date_row           t_date_table;
   v_number_row         t_number_table;
   v_char_dummy         varchar2(1);
   v_date_dummy         date;
   v_number_dummy       number;
   v_prev_row           APEX_APPLICATION_GLOBAL.VC_ARR2;
   v_columns            APEX_APPLICATION_GLOBAL.VC_ARR2;
   v_current_row        number default 0;
   v_desc_tab           DBMS_SQL.DESC_TAB2;
   v_inside             boolean default false;
   v_sql                largevarchar2;
   v_bind_variables     DBMS_SQL.VARCHAR2_TABLE;
   v_buffer             largevarchar2;
   v_bind_var_name      varchar2(255);
   v_binded             boolean; 
   v_format_mask        formatmask;  
   v_rowset_count number := 0; 

    v_column_alias    APEX_APPLICATION_PAGE_IR_COL.column_alias%TYPE;
    v_row_color       varchar2(10); 
    v_row_back_color  varchar2(10);
    v_cell_color      varchar2(10);
    v_cell_back_color VARCHAR2(10);     
    v_column_type     columntype;
    v_cell_data       t_cell_data;    
    v_i number := 0;

    v_strings          apex_application_global.vc_arr2;
    v_rownum           binary_integer default 1;
    
    v_style_id         binary_integer;         
    v_string_buffer    t_large_varchar2; 
    v_break_header     t_large_varchar2;
    v_last_agg_obj     APEX_APPLICATION_GLOBAL.VC_ARR2;
    a_col_name_plus_width  apex_application_global.vc_arr2;    
    v_col_width            number;
    v_coefficient          number;
    
    procedure print_char_cell(p_coll      in binary_integer, 
                              p_row       in binary_integer, 
                              p_string    in varchar2,
                              p_clob      in out nocopy CLOB,
                              p_buffer    IN OUT NOCOPY VARCHAR2,
                              p_style_id  in number default null
                             )
    is
     v_style varchar2(20);
    begin
      if p_style_id is not null then
       v_style := ' s="'||p_style_id||'" ';
      end if;
      add(p_clob,p_buffer,'<c r="'||get_cell_name(p_coll,p_row)||'" t="s" '||v_style||'>'||chr(10)
                         ||'<v>' || to_char(v_strings.count)|| '</v>'||chr(10)                 
                         ||'</c>'||chr(10));
      v_strings(v_strings.count + 1) := p_string;
    end print_char_cell;
    --
    procedure print_number_cell(p_coll      in binary_integer, 
                                p_row       in binary_integer, 
                                p_value     in varchar2,
                                p_clob      in out nocopy clob,
                                p_buffer    IN OUT NOCOPY VARCHAR2,
                                p_style_id  in number default null
                               )
    is
      v_style varchar2(20);
    begin
      if p_style_id is not null then
       v_style := ' s="'||p_style_id||'" ';
      end if;
      add(p_clob,p_buffer,'<c r="'||get_cell_name(p_coll,p_row)||'" '||v_style||'>'||chr(10)
                         ||'<v>'||p_value|| '</v>'||chr(10)
                         ||'</c>'||chr(10));
    
    end print_number_cell; 
    --
    procedure print_agg(p_agg_obj APEX_APPLICATION_GLOBAL.VC_ARR2,
                        p_rownum    in out binary_integer,
                        p_clob      in out nocopy clob,
                        p_buffer    IN OUT NOCOPY VARCHAR2)
    is
        v_agg_clob         clob;
        v_agg_buffer       t_large_varchar2;
        v_agg_strings_cnt  binary_integer default 1; 
    begin
        dbms_lob.createtemporary(v_agg_clob,true);
        /*PRINTAGG*/
        if p_agg_obj.last IS NOT NULL THEN
            DBMS_LOB.TRIM(v_agg_clob,0);
            v_agg_buffer := '';       
            v_agg_strings_cnt := 1;       
            
            FOR y in p_agg_obj.FIRST..p_agg_obj.last
            LOOP
                v_agg_strings_cnt := greatest(length(regexp_replace('[^:]','')) + 1,v_agg_strings_cnt);
                v_style_id := get_aggregate_style_id(p_font         => '',
                                                    p_back         => '',
                                                    p_data_type    => 'CHAR',
                                                    p_format_mask  => '');
                print_char_cell(p_coll       => y,
                                p_row        => v_rownum,
                                p_string     => get_xmlval( rtrim( p_agg_obj(y),chr(10))),
                                p_clob       => v_agg_clob,
                                p_buffer    =>  v_agg_buffer,           
                                p_style_id   => v_style_id);        
            END LOOP;
            add(p_clob,v_buffer,'<row ht="'||(v_agg_strings_cnt * STRING_HEIGHT)||'">'||chr(10),TRUE);
            add(v_agg_clob,v_agg_buffer,' ',TRUE);
            dbms_lob.copy( dest_lob => p_clob,
                            src_lob => v_agg_clob,
                            amount => dbms_lob.getlength(v_agg_clob),
                            dest_offset => dbms_lob.getlength(p_clob),
                            src_offset => 1);
            add(p_clob,v_buffer,'</row>'||chr(10));
            p_rownum := p_rownum + 1;
        end if;
        dbms_lob.freetemporary(v_agg_clob);
  end print_agg;
  --
  begin    
    pragma inline(add,'YES');     
    
    add(v_clob,v_buffer,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'||chr(10));
    add(v_clob,v_buffer,'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <dimension ref="A1"/>
        <sheetViews>
            <sheetView tabSelected="1" workbookViewId="0">
            <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
            <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
            </sheetView>
            </sheetViews>
        <sheetFormatPr baseColWidth="10" defaultColWidth="10" defaultRowHeight="15"/>'||chr(10),TRUE);
    
    a_col_name_plus_width := APEX_UTIL.STRING_TO_TABLE(rtrim(p_width_str,','),','); 
    
    v_coefficient := nvl(p_coefficient,WIDTH_COEFFICIENT);
    if v_coefficient =  0 then
      v_coefficient := WIDTH_COEFFICIENT;
    end if;     
    add(v_clob,v_buffer,'<cols>'||chr(10));
    <<column_widths>>
    for i in 1..l_report.displayed_columns.count   loop
       v_column_alias := get_column_alias(i);
      -- if current column is not control break column
      if apex_plugin_util.get_position_in_list(l_report.break_on,v_column_alias) is null then 
        v_i := v_i + 1;
        begin                    
            v_col_width  := round(to_number(a_col_name_plus_width(v_i))/ v_coefficient);
          exception
             when others then             
                v_col_width := -1;  
          end;
          
          if v_col_width >= 0 then
            add(v_clob,v_buffer,'<col min="'||v_i||'" max="'||v_i||'" width="'||v_col_width||'" customWidth="1" />'||chr(10));        
          else
            add(v_clob,v_buffer,'<col min="'||v_i||'" max="'||v_i||'" width="10" customWidth="0" />'||chr(10));        
          end if;
      end if;  
    end loop column_widths;
    v_i := 0;
    add(v_clob,v_buffer,'</cols>'||chr(10));
     --!
    add(v_clob,v_buffer,'<sheetData>'||chr(10));
     --column header
    add(v_clob,v_buffer,'<row>'||chr(10));      

    v_cur := dbms_sql.open_cursor(2); 
    v_sql := apex_plugin_util.replace_substitutions(p_value  => l_report.report.sql_query,
                                                    p_escape => false);    
    dbms_sql.parse(v_cur,v_sql,dbms_sql.native);     
    dbms_sql.describe_columns2(v_cur,v_colls_count,v_desc_tab);    
    --skip internal primary key if need
    for i in 1..v_desc_tab.count loop      
      if lower(v_desc_tab(i).col_name) = 'apxws_row_pk' then
        l_report.skipped_columns := 1;
      end if;
    end loop;
    
    l_report.start_with := 1 + 
                           l_report.skipped_columns +
                           nvl(l_report.break_really_on.count,0) + 
                           l_report.row_highlight.count + 
                           l_report.col_highlight.count;
    l_report.end_with   := l_report.skipped_columns + 
                           nvl(l_report.break_really_on.count,0) + 
                           l_report.displayed_columns.count  + 
                           l_report.row_highlight.count + 
                           l_report.col_highlight.count;    
                           
    log('l_report.start_with='||l_report.start_with);
    log('l_report.end_with='||l_report.end_with);
    log('l_report.skipped_columns='||l_report.skipped_columns);
    -- init column datatypes and format masks
    for i in 1..v_desc_tab.count loop      
      log('column_type='||v_desc_tab(i).col_type);
      -- don't know why column types in dbms_sql.describe_columns2 do not correspond to types in dbms_types      
      if v_desc_tab(i).col_type in (dbms_types.TYPECODE_TIMESTAMP_TZ,181,
                                    dbms_types.TYPECODE_TIMESTAMP_LTZ,231,
                                    dbms_types.TYPECODE_TIMESTAMP,180,
                                    dbms_types.TYPECODE_DATE) 
      then
         prepare_col_format_mask(p_col_number => i,p_data_type => v_desc_tab(i).col_type);      
      elsif v_desc_tab(i).col_type = dbms_types.TYPECODE_NUMBER then
         l_report.column_types(i) := 'NUMBER';
      else
         l_report.column_types(i) := 'STRING';
         log('column_type='||v_desc_tab(i).col_type||' STRING');
      end if;        
    end loop;
    
    v_bind_variables := wwv_flow_utilities.get_binds(v_sql);
   
    <<headers>>
    for i in 1..l_report.displayed_columns.count   loop
       v_column_alias := get_column_alias(i);
      -- if current column is not control break column
      if apex_plugin_util.get_position_in_list(l_report.break_on,v_column_alias) is null then       
       /* 
        h_cell.column_alias := v_column_alias;
        h_cell.header_align :=  get_header_alignment(v_column_alias);
        h_cell.align :=  get_column_alignment(v_column_alias);
        h_cell.column_type := get_column_types(i);
        h_cell.format_mask := get_col_format_mask(v_column_alias);
        h_cell.column_text := get_column_names(v_column_alias);*/        
        print_char_cell( p_coll      => i, 
                        p_row       => v_rownum, 
                        p_string    =>   get_xmlval(get_column_names(v_column_alias)),
                        p_clob      => v_clob,
                        p_buffer    => v_buffer,
                        p_style_id  => get_header_style_id(p_align  => get_header_alignment(v_column_alias))
                      );        
      end if;  
    end loop headers;    
    v_rownum := v_rownum + 1;
    add(v_clob,v_buffer,'</row>'||chr(10)); 

    log('<<bind variables>>');
    
    <<bind_variables>>
    for i in 1..v_bind_variables.count loop      
      v_bind_var_name := ltrim(v_bind_variables(i),':');
      if v_bind_var_name = 'APXWS_MAX_ROW_CNT' then      
         -- remove max_rows
         DBMS_SQL.BIND_VARIABLE (v_cur,'APXWS_MAX_ROW_CNT',p_max_rows);    
         log('Bind variable ('||i||')'||v_bind_var_name||'<'||p_max_rows||'>');
      else
        v_binded := false; 
        --first look report bind variables (filtering, search etc)
        <<bind_report_variables>>
        for a in 1..l_report.report.binds.count loop
          if v_bind_var_name = l_report.report.binds(a).name then
             DBMS_SQL.BIND_VARIABLE (v_cur,v_bind_var_name,l_report.report.binds(a).value);
             log('Bind variable as report variable ('||i||')'||v_bind_var_name||'<'||l_report.report.binds(a).value||'>');
             v_binded := true;
             exit;
          end if;
        end loop bind_report_variables;
        -- substantive strings in sql-queries can have bind variables too
        -- these variables are not in v_report.binds
        -- and need to be binded separately
        if not v_binded then
          DBMS_SQL.BIND_VARIABLE (v_cur,v_bind_var_name,v(v_bind_var_name));
          log('Bind variable ('||i||')'||v_bind_var_name||'<'||v(v_bind_var_name)||'>');          
        end if;        
       end if; 
    end loop;          
    
    log('<<define_columns>>');    
    for i in 1..v_colls_count loop
       log('define column '||i);
       log('column type '||v_desc_tab(i).col_type);      
       if l_report.column_types(i) = 'DATE' then   
         dbms_sql.define_column(v_cur, i, v_date_dummy);
       elsif l_report.column_types(i) = 'NUMBER' then   
         dbms_sql.define_column(v_cur, i, v_number_dummy);
       else --STRING
         dbms_sql.define_column(v_cur, i, v_char_dummy,32767);
       end if;         
    end loop define_columns;    
    
    v_result := dbms_sql.execute(v_cur);         
   
    log('<<main_cycle>>');
    <<main_cycle>>
    LOOP 
         IF DBMS_SQL.FETCH_ROWS(v_cur)>0 THEN          
         log('<<fetch>>');
           -- get column values of the row 
            v_current_row := v_current_row + 1;
            <<query_columns>>
            for i in 1..v_colls_count loop               
               log('column type '||v_desc_tab(i).col_type);
               v_row(i) := ' ';
               v_date_row(i) := NULL;
               v_number_row(i) := NULL;
               if l_report.column_types(i) = 'DATE' then
                 dbms_sql.column_value(v_cur, i,v_date_row(i));
                 v_row(i) := to_char(v_date_row(i));
               elsif l_report.column_types(i) = 'NUMBER' then
                dbms_sql.column_value(v_cur, i,v_number_row(i));
                v_row(i) := to_char(v_number_row(i));
               else 
                 dbms_sql.column_value(v_cur, i,v_row(i));                 
               end if;  
            end loop query_columns;     
            --check control break
            if v_current_row > 1 then
             if is_control_break(v_row,v_prev_row) then
                v_inside := false;                                                         
                print_agg(v_last_agg_obj,v_rownum,v_clob,v_buffer);                
             end if;
            end if;
            if not v_inside then
                v_break_header :=  print_control_break_header_obj(v_row);
                if v_break_header IS NOT NULL then
                    add(v_clob,v_buffer,'<row>'||chr(10));
                    print_char_cell(p_coll      => 1,
                                    p_row       => v_rownum,
                                    p_string    =>  get_xmlval(v_break_header),
                                    p_clob      => v_clob,
                                    p_buffer    => v_buffer,
                                    p_style_id  => get_header_style_id(p_back => NULL,p_align  => 'left')
                                    );         
                    v_rownum := v_rownum + 1;
                    add(v_clob,v_buffer,'</row>'||chr(10));
                end if;
                              
               v_last_agg_obj := get_aggregate(v_row);                           
               v_inside := true;
            END IF;            --                        
            for i in 1..v_colls_count loop
              v_prev_row(i) := v_row(i);                           
            end loop;
            
            add(v_clob,v_buffer,'<row>'||chr(10));
            /* CELLS INSIDE ROW PRINTING*/
            v_row_color := NULL;
            v_row_back_color  := NULL;
            <<row_highlights>>
            for h in 1..l_report.row_highlight.count loop
            begin 
                -- J.P.Lourens 9-Oct-16 
                -- current_row is based on report_sql which starts with the highlight columns, then the skipped columns and then the rest
                -- So to capture the highlight values the value for l_report.skipped_columns should NOT be taken into account
                if get_current_row(v_row,/*l_report.skipped_columns + */l_report.row_highlight(h).cond_number) IS NOT NULL THEN
                    v_row_color       := l_report.row_highlight(h).highlight_row_font_color;
                    v_row_back_color  := l_report.row_highlight(h).highlight_row_color;
                end if;
                exception       
                  when no_data_found then
                      log('row_highlights: ='||' end_with='||l_report.end_with||' agg_cols_cnt='||l_report.agg_cols_cnt||' COND_NUMBER='||l_report.row_highlight(h).cond_number||' h='||h);
                end; 
            end loop row_highlights;
            
            
            <<visible_columns>>
            v_i := 0;
            for i in l_report.start_with..l_report.end_with loop
                v_i := v_i + 1;
                v_cell_color       := NULL;
                v_cell_back_color  := NULL;
                v_cell_data.value  := NULL;  
                v_cell_data.text   := NULL; 
                v_column_alias     := get_column_alias_sql(i);
                v_column_type      := get_column_types(i);
                v_format_mask      := get_col_format_mask(v_column_alias);
                
                IF v_column_type = 'DATE' THEN
                    v_cell_data := get_cell(get_current_row(v_row,i),v_format_mask,get_current_row(v_date_row,i));
                ELSIF  v_column_type = 'NUMBER' THEN      
                    v_cell_data := get_cell(get_current_row(v_row,i),v_format_mask,get_current_row(v_number_row,i));
                ELSE --STRING
                    v_format_mask := NULL;
                    v_cell_data.VALUE  := NULL;  
                    v_cell_data.datatype := 'STRING';
                    v_cell_data.text   := get_current_row(v_row,i);
                end if; 
                
                --check that cell need to be highlighted
                <<cell_highlights>>
                for h in 1..l_report.col_highlight.count loop
                    begin
                    -- J.P.Lourens 9-Oct-16 
                    -- current_row is based on report_sql which starts with the highlight columns, then the skipped columns and then the rest
                    -- So to capture the highlight values the value for l_report.skipped_columns should NOT be taken into account
                    if get_current_row(v_row,/*l_report.skipped_columns + */l_report.col_highlight(h).cond_number) is not null 
                        and v_column_alias = l_report.col_highlight(h).condition_column_name 
                    then
                        v_cell_color       := l_report.col_highlight(h).highlight_cell_font_color;
                        v_cell_back_color  := l_report.col_highlight(h).highlight_cell_color;
                    end if;
                    exception
                when no_data_found then
                    log('col_highlights: ='||' end_with='||l_report.end_with||' agg_cols_cnt='||l_report.agg_cols_cnt||' COND_NUMBER='||l_report.col_highlight(h).cond_number||' h='||h); 
                end;
                END loop cell_highlights;
            
                begin            
                    v_style_id := get_style_id(p_font          => nvl(v_cell_color,v_row_color),
                                                p_back         => nvl(v_cell_back_color,v_row_back_color),
                                                p_data_type    => v_cell_data.datatype,
                                                p_format_mask  => v_format_mask,
                                                p_align        => lower(get_column_alignment(v_column_alias))
                                                );
                    if v_cell_data.datatype in ('NUMBER') then
                        print_number_cell(p_coll        => v_i, 
                                            p_row       => v_rownum, 
                                            p_value     => v_cell_data.value,
                                            p_clob      => v_clob,
                                            p_buffer    => v_buffer,
                                            p_style_id  => v_style_id
                                        );
                        
                    elsif  v_cell_data.datatype in ('DATE') then
                        add(v_clob,v_buffer,'<c r="'||get_cell_name(v_i,v_rownum)||'"  s="'||v_style_id||'">'||chr(10)
                                            ||'<v>'|| v_cell_data.value|| '</v>'||chr(10)
                                            ||'</c>'||chr(10)
                            );
                    else --STRING
                        print_char_cell(p_coll        => v_i,
                                        p_row         => v_rownum,
                                        p_string      => get_xmlval(v_cell_data.text),
                                        p_clob        => v_clob,
                                        p_buffer      => v_buffer,
                                        p_style_id    => v_style_id
                                        );
                    END IF;       
                exception
                WHEN no_data_found THEN
                    null;
                end;
            
            end loop visible_columns;            

            add(v_clob,v_buffer,'</row>'||chr(10));         
            v_rownum := v_rownum + 1;
         ELSE
           EXIT; 
         END IF;
    END LOOP main_cycle;    
    if v_inside then       
       v_inside := false;
       print_agg(v_last_agg_obj,v_rownum,v_clob,v_buffer);
    end if;  
    dbms_sql.close_cursor(v_cur);  

    add(v_clob,v_buffer,'</sheetData>'||chr(10));
    if p_autofilter = 'Y' then
        add(v_clob,v_buffer,'<autoFilter ref="A1:' || get_cell_name(v_colls_count,v_rownum-1) || '"/>');
    end if;
    add(v_clob,v_buffer,'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'||chr(10),TRUE);
    
    add(v_strings_clob,v_string_buffer,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'||chr(10));
    add(v_strings_clob,v_string_buffer,'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' || v_strings.count() || '" uniqueCount="' || v_strings.count() || '">'||chr(10));
    
    for i in 1 .. v_strings.count() loop
    add(v_strings_clob,v_string_buffer,'<si><t>'||dbms_xmlgen.convert( substr( v_strings( i ), 1, 32000 ) ) || '</t></si>'||chr(10));
    end loop; 
    add(v_strings_clob,v_string_buffer,'</sst>'||chr(10),TRUE);
  end generate_from_report;
  ------------------------------------------------------------------------------
    
   procedure get_report(p_app_id          IN NUMBER,
                        p_page_id         IN NUMBER,  
                        p_region_id       IN NUMBER,
                        p_max_rows        IN NUMBER,            -- maximum rows for export   
                        v_clob in out nocopy CLOB,
                        p_strings in out nocopy CLOB,
                        p_col_length IN VARCHAR2,
                        p_width_coefficient IN number, 
                        p_autofilter IN CHAR DEFAULT 'Y'      
                       )                 
  is  
  begin
    dbms_lob.trim (v_debug,0);
  
    log('p_app_id='||p_app_id);
    log('p_page_id='||p_page_id);
    log('p_region_id='||p_region_id);
    log('p_max_rows='||p_max_rows);
    log('p_col_length='||p_col_length);
    log('p_autofilter='||p_autofilter);    
    
    init_t_report(p_app_id,p_page_id,p_region_id);
    generate_from_report(v_clob,p_strings,p_col_length,p_width_coefficient,p_max_rows,p_autofilter);
  end get_report;   
  
  ------------------------------------------------------------------------------
  /* 
    function to handle cases of 'in' and 'not in' conditions for highlights
       used in cursor cur_highlight
    
    Author: Srihari Ravva
  */ 
  function get_highlight_in_cond_sql(p_condition_expression  in APEX_APPLICATION_PAGE_IR_COND.CONDITION_EXPRESSION%TYPE,
                                     p_condition_sql         in APEX_APPLICATION_PAGE_IR_COND.CONDITION_SQL%TYPE,
                                     p_condition_column_name in APEX_APPLICATION_PAGE_IR_COND.CONDITION_COLUMN_NAME%TYPE)
  return varchar2 
  is
    v_condition_sql_tmp     varchar2(32767);
      v_condition_sql            varchar2(32767);
      v_arr_cond_expr APEX_APPLICATION_GLOBAL.VC_ARR2;
      v_arr_cond_sql APEX_APPLICATION_GLOBAL.VC_ARR2;    
  begin
    v_condition_sql := REPLACE(REPLACE(p_condition_sql,'#APXWS_HL_ID#','1'),'#APXWS_CC_EXPR#','"'||p_condition_column_name||'"');
      v_condition_sql_tmp := SUBSTR(v_condition_sql,INSTR(v_condition_sql,'#'),INSTR(v_condition_sql,'#',-1)-INSTR(v_condition_sql,'#')+1);
    
    v_arr_cond_expr := APEX_UTIL.STRING_TO_TABLE(p_condition_expression,',');
    v_arr_cond_sql := APEX_UTIL.STRING_TO_TABLE(v_condition_sql_tmp,',');
    
    for i in 1..v_arr_cond_expr.count
    loop
        -- consider everything as varchar2
        -- 'in' and 'not in' highlight conditions are not possible for DATE columns from IR
        v_condition_sql := REPLACE(v_condition_sql,v_arr_cond_sql(i),''''||TO_CHAR(v_arr_cond_expr(i))||'''');
    end loop;
    return v_condition_sql;
  end get_highlight_in_cond_sql;  

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
  function get_max_rows (p_app_id      in number,
                         p_page_id     in number,
                         p_region_id   IN NUMBER)
  return number
  is 
    v_max_row_count number;
  begin
    select max_row_count 
    into v_max_row_count
    from APEX_APPLICATION_PAGE_IR
    where application_id = p_app_id
      and page_id = p_page_id
      and region_id = p_region_id;
       
     return v_max_row_count;
  end get_max_rows;   
  ------------------------------------------------------------------------------  
  function get_file_name (p_app_id      IN NUMBER,
                          p_page_id     IN NUMBER,
                          p_region_id   IN NUMBER)
  return varchar2
  is 
    v_filename varchar2(255);
  begin
    select filename 
    into v_filename
    from APEX_APPLICATION_PAGE_IR
    where application_id = p_app_id
      and page_id = p_page_id
      and region_id = p_region_id;
     
     return apex_plugin_util.replace_substitutions(nvl(v_filename,'Excel'));
  end get_file_name;   
  ------------------------------------------------------------------------------
  -- download binary file
  procedure download(p_data        IN OUT NOCOPY BLOB,
                     p_mime_type   IN VARCHAR2,
                     p_file_name   IN VARCHAR2)
  is
  begin
    owa_util.mime_header( ccontent_type=> p_mime_type, 
                          bclose_header => false );
    sys.htp.p('Content-Length: ' || dbms_lob.getlength( p_data ) );
    sys.htp.p('Content-disposition: attachment; filename='||p_file_name );    
    sys.htp.p('Cache-Control: must-revalidate, max-age=0');
    sys.htp.p('Expires: Thu, 01 Jan 1970 01:00:00 CET');
    sys.htp.p('Set-Cookie: GPV_DOWNLOAD_STARTED=1;');
    owa_util.http_header_close;    
    wpg_docload.download_file( p_data );  
  exception
     when others then 
        raise_application_error(-20001,'Download (blob)'||SQLERRM);
  end download;
  ------------------------------------------------------------------------------
  
  procedure download_debug(p_app_id      IN NUMBER,
                           p_page_id     IN NUMBER,
                           p_region_id   IN NUMBER, 
                           p_col_length  IN VARCHAR2 DEFAULT NULL,
                           p_max_rows    IN NUMBER,
                           p_autofilter  IN CHAR DEFAULT 'Y'                           
                          )
  is
    v_blob     blob;
    v_cells    clob;
    v_strings  clob;
    v_desc_offset PLS_INTEGER := 1;
    v_src_offset  PLS_INTEGER := 1;
    v_lang        PLS_INTEGER := 0;
    v_warning     PLS_INTEGER := 0;   
  begin
    dbms_lob.createtemporary(v_cells,true);
    dbms_lob.createtemporary(v_strings,true);
    dbms_lob.createtemporary(v_blob,true);
    
    get_report(p_app_id,
               p_page_id ,  
               p_region_id,
               p_max_rows,            -- maximum rows for export   
               v_cells ,
               v_strings,
               p_col_length,
               WIDTH_COEFFICIENT,
               p_autofilter);     

    log('LogFinish',TRUE);
    dbms_lob.converttoblob(v_blob, v_debug, dbms_lob.getlength(v_debug), v_desc_offset, v_src_offset, dbms_lob.default_csid, v_lang, v_warning);
    sys.htp.init;
    sys.owa_util.mime_header('text/txt', FALSE );
    sys.htp.p('Content-length: ' || sys.dbms_lob.getlength( v_blob));
    sys.htp.p('Content-Disposition: attachment; filename="log.txt"' );
    sys.htp.p('Cache-Control: must-revalidate, max-age=0');
    sys.htp.p('Expires: Thu, 01 Jan 1970 01:00:00 CET');
    sys.htp.p('Set-Cookie: GPV_DOWNLOAD_STARTED=1;');
    sys.owa_util.http_header_close;
    sys.wpg_docload.download_file( v_blob );

    dbms_lob.freetemporary(v_blob);
  end download_debug;

  ------------------------------------------------------------------------------
  
  procedure download_excel(p_app_id      IN NUMBER,
                           p_page_id     IN NUMBER,
                           p_region_id   IN NUMBER, 
                           p_col_length  IN VARCHAR2 DEFAULT NULL,
                           p_max_rows    IN NUMBER,
                           p_autofilter  IN CHAR DEFAULT 'Y'                           
                          )
  is
    t_template blob;
    t_excel    blob;
    v_cells    clob;
    v_strings  clob;
    zip_files  as_zip.file_list;
  begin        
    dbms_lob.createtemporary(t_excel,true);    
    dbms_lob.createtemporary(v_cells,true);
    dbms_lob.createtemporary(v_strings,true);

    get_report(p_app_id,
               p_page_id ,  
               p_region_id,
               p_max_rows,            -- maximum rows for export   
               v_cells ,
               v_strings,
               p_col_length,
               WIDTH_COEFFICIENT,
               p_autofilter        
              );
              
    select file_content
    into t_template
    from apex_appl_plugin_files 
    where file_name = 'ExcelTemplate.zip'
      and application_id = p_app_id
      and plugin_name='AT.FRT.GPV_IR_TO_MSEXCEL';
    
    zip_files  := as_zip.get_file_list( t_template );
    for i in zip_files.first() .. zip_files.last loop
      as_zip.add1file( t_excel, zip_files( i ), as_zip.get_file( t_template, zip_files( i ) ) );
    end loop;    
      
    add1file( t_excel, 'xl/styles.xml', get_styles_xml);    
    add1file( t_excel, 'xl/worksheets/Sheet1.xml', v_cells);
    add1file( t_excel, 'xl/sharedStrings.xml',v_strings);
    add1file( t_excel, 'xl/_rels/workbook.xml.rels',t_sheet_rels);    
    add1file( t_excel, 'xl/workbook.xml',t_workbook);    
    
    as_zip.finish_zip( t_excel );
 
    download(p_data      => t_excel,
             p_mime_type => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
             p_file_name => get_file_name (p_app_id,p_page_id,p_region_id)||'.xlsx;'
            );
    dbms_lob.freetemporary(t_excel);
    dbms_lob.freetemporary(v_cells);
    dbms_lob.freetemporary(v_strings);    
  end download_excel;
  
begin
  dbms_lob.createtemporary(v_debug,true, DBMS_LOB.CALL);  
END IR_TO_XLSX;
/

create or replace PACKAGE  "IR_TO_MSEXCEL" 
  AUTHID CURRENT_USER
as
  FUNCTION render  (p_dynamic_action in apex_plugin.t_dynamic_action,
                    p_plugin         in apex_plugin.t_plugin )
  return apex_plugin.t_dynamic_action_render_result; 
  
  function ajax (p_dynamic_action in apex_plugin.t_dynamic_action,
                 p_plugin         in apex_plugin.t_plugin )
  return apex_plugin.t_dynamic_action_ajax_result;

  function is_ir2msexcel 
  return boolean;

end IR_TO_MSEXCEL;
/

create or replace PACKAGE BODY  "IR_TO_MSEXCEL" 
as
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
        and aapda.page_id = nv('APP_PAGE_ID') 
        and aapda.application_id = nv('APP_ID')
        and r.page_id = nv('APP_PAGE_ID') 
        and r.application_id = nv('APP_ID'); 
      
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
                  where r.page_id = nv('APP_PAGE_ID')
                    and r.application_id =nv('APP_ID')
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
	p_autofilter         char;
    v_maximum_rows       number;
    v_dummy              apex_plugin.t_dynamic_action_ajax_result;
    v_affected_region_id apex_application_page_da_acts.affected_region_id%type;
  begin      
      v_plugin_running := true;
      p_download_type := nvl(p_dynamic_action.attribute_02,'E');
  	  p_autofilter:= nvl(p_dynamic_action.attribute_04,'Y');
      v_affected_region_id := get_affected_region_id(p_dynamic_action_id => p_dynamic_action.ID
                                                    ,p_html_region_id    => apex_application.g_x03);
      
      v_maximum_rows := nvl(nvl(p_dynamic_action.attribute_01,
                                IR_TO_XLSX.get_max_rows (p_app_id    => apex_application.g_x01,
                                                         p_page_id   => apex_application.g_x02,
                                                         p_region_id => v_affected_region_id)
                                ),1000);                                               
      if p_download_type = 'E' then 
        ir_to_xlsx.download_excel(p_app_id        => apex_application.g_x01,
                                  p_page_id      => apex_application.g_x02,
                                  p_region_id    => v_affected_region_id,
                                  p_col_length   => apex_application.g_x04||p_custom_width,
                                  p_max_rows     => v_maximum_rows,
                                  p_autofilter   => p_autofilter
                                  );
      elsif p_download_type = 'T' then 
        ir_to_xlsx.download_debug(p_app_id        => apex_application.g_x01,
                                  p_page_id      => apex_application.g_x02,
                                  p_region_id    => v_affected_region_id,
                                  p_col_length   => apex_application.g_x04||p_custom_width,
                                  p_max_rows     => v_maximum_rows,
                                  p_autofilter   => p_autofilter
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

