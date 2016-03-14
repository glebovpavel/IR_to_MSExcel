function getColWidthsDelimeteredString(p_region_static_id) {  
    var colWidthsDelimeteredString = "";    
    var region_selector;
    if(p_region_static_id)
        {
            region_selector = '#' + p_region_static_id + '_data_panel';
        }
    $('div' + region_selector + ' table.a-IRR-table th:not(".a-IRR-header--group")').parent().first().find('th').each(function( index,elmt ){    
      colWidthsDelimeteredString = colWidthsDelimeteredString + $(elmt).width() + "\,";  
    });
    return colWidthsDelimeteredString;
}

function addDownloadXLSXIcon(plugin_id_in,p_region_static_id) {     
  $('body').on( "dialogopen", function( event, ui ) {
     if($('span.ui-dialog-title').text() === apex.lang.getMessage( "APEXIR_DOWNLOAD")) {
        var $dialog_window = $(event.target);
        var current_region_id = $dialog_window.attr('id').match(/(.+)_dialog_js/)[1]; 
        var re = new RegExp(p_region_static_id, 'g'); 
         
        if(!p_region_static_id || current_region_id.match(re) != null){        
          $dialog_window.find('table.a-IRR-dialogTable tbody td a').closest('tr').append('<tr><td nowrap="nowrap"><a href="javascript:get_excel_gpv(' + "'" + current_region_id + "','" + plugin_id_in +  "'" +')"><img src="/i/ws/download_xls_64x64.gif" alt="XLSX" title="XLSX"></a></td></tr>');
          $dialog_window.find('table.a-IRR-dialogTable tbody td span').closest('tr').append( '<tr><td align="center" nowrap="nowrap"><span>XLSX</span></td></tr>');
        }    
     } 
  }); 
}

function get_excel_gpv(p_region_static_id,plugin_id_in) {
  apex.navigation.redirect(
        "wwv_flow.show?p_flow_id=" + $('#pFlowId').val() 
                   + "&p_flow_step_id=" + $('#pFlowStepId').val() 
                   + "&p_instance="+ $v('pInstance')  
                   + "&p_request=" + "PLUGIN=" + plugin_id_in
                   + "&x01=" + $('#pFlowId').val()
                   + "&x02=" + $('#pFlowStepId').val()
                   + "&x03=" + p_region_static_id
                   + "&x04=" + getColWidthsDelimeteredString(p_region_static_id)
   );     
}
