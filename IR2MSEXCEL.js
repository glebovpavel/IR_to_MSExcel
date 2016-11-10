(function(parent, $, undefined){
  
	if (parent.excel_gpv === undefined) {

		parent.excel_gpv = function () {
      //private
      function getColWidthsDelimeteredString ( p_region_static_id ) {
        if ( $.type(p_region_static_id) !== "string"){
          throw("Expecting a string as static ID (arugment: p_region_static_id)");
        };

        var colWidthsDelimeteredString = "";
        var region_selector = '#' + p_region_static_id + '_data_panel';

        $('div' + region_selector + ' table.a-IRR-table th:not(".a-IRR-header--group")').parent().first().find('th').each(function( index,elmt ){
          colWidthsDelimeteredString = colWidthsDelimeteredString + $(elmt).width() + "\,";
        });
        return colWidthsDelimeteredString;
      }

      //exposed
      function getExcel ( p_region_static_id, plugin_id_in ) {
        apex.navigation.redirect(
              "wwv_flow.show?p_flow_id=" + $v('pFlowId')
                         + "&p_flow_step_id=" + $v('pFlowStepId')
                         + "&p_instance="+ $v('pInstance')
                         + "&p_request=" + "PLUGIN=" + plugin_id_in
                         + "&p_debug=" + $v("pdebug")
                         + "&x01=" + $v('pFlowId')
                         + "&x02=" + $v('pFlowStepId')
                         + "&x03=" + p_region_static_id
                         + "&x04=" + getColWidthsDelimeteredString(p_region_static_id)
         );
      }
      
      function addDownloadXLSXIcon ( plugin_id_in, p_region_static_id ) {
        $('body').on( "dialogopen", function( event, ui ) {
          if($('span.ui-dialog-title').text() === apex.lang.getMessage( "APEXIR_DOWNLOAD")) {
            var $dialog_window = $(event.target);
            var current_region_id = $dialog_window.attr('id').match(/(.+)_dialog_js/)[1];
            var re = new RegExp(p_region_static_id, 'g');
            
            var html = apex.util.htmlBuilder();
            html.markup('<tr>')
                .markup('<td')
                  .attr('nowrap', 'nowrap')
                .markup('>')
                .markup('<a')
                  .attr('href','javascript:excel_gpv.getExcel(' + "'" + current_region_id + "','" + plugin_id_in +  "'" +');')
                .markup('>')
                .markup('<img')
                  .attr('src', apex_img_dir+'ws/download_xls_64x64.gif')
                  .attr('alt', 'XLSX')
                  .attr('title', 'XLSX')
                .markup(' />')
                .markup('</a>')
                .markup('</td>')
                .markup('</tr>');
            
            if(!p_region_static_id || current_region_id.match(re) != null){            
              $dialog_window.find('table.a-IRR-dialogTable tbody td a')
                .closest('tr')
                .append( html.toString() );
                
              $dialog_window.find('table.a-IRR-dialogTable tbody td span')
                .closest('tr')
                .append( '<tr><td align="center" nowrap="nowrap"><span>XLSX</span></td></tr>');
            };
          }
        });
      }
      
      return {
        getExcel:getExcel
      , addDownloadXLSXIcon: addDownloadXLSXIcon
      };
  
		}();
	}
})(window, apex.jQuery)