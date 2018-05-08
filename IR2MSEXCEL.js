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

      function blockUIForDownload() {
       var token = new Date().getTime(); //use the current timestamp as the token value
       var fileDownloadCheckTimer;
       var mySpinner;
       var mySpinner = apex.widget.waitPopup();
       fileDownloadCheckTimer = window.setInterval(function () {
         var cookieValue = apex.storage.getCookie('GPV_DOWNLOAD_STARTED');
         console.log('Check'); 
         if (cookieValue)
          {
           window.clearInterval(fileDownloadCheckTimer);
           document.cookie = "GPV_DOWNLOAD_STARTED=; expires=Thu, 01 Jan 1970 00:00:00 GMT"; //clears this cookie value
           mySpinner.remove();
           $(".gpvCloseButton").click(); //close download window
          }
       }, 1000);
     }

      //exposed
      function getExcel ( p_region_static_id, plugin_id_in ) {
        blockUIForDownload();
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
      
      function addDownloadXLSXIcon ( plugin_id_in, p_region_static_id,p_version) {
        $('body').on( "dialogopen", function( event, ui ) {
          if($('span.ui-dialog-title').text() === apex.lang.getMessage( "APEXIR_DOWNLOAD")) {
            var $dialog_window = $(event.target);
            var current_region_id = $dialog_window.attr('id').match(/(.+)_dialog_js/)[1];
            var re = new RegExp(p_region_static_id, 'g');

            if (p_version > 5.1) 
             {
                var html = apex.util.htmlBuilder();
                html.markup('<li')
                    .attr('class', 'a-IRR-iconList-item')
                    .markup('>')                    
                    .markup('<a')
                    .attr('class', 'a-IRR-iconList-link')
                    .attr('href','javascript:excel_gpv.getExcel(' + "'" + current_region_id + "','" + plugin_id_in +  "'" +');')
                    .attr('id', 'download_excel_gpv')
                    .markup('>')
                    .markup('<span')
                    .attr('class','a-IRR-iconList-icon a-Icon icon-irr-dl-xls')
                    .markup('>')
                    .markup('</span>')
                    .markup('<span')
                    .attr('class','a-IRR-iconList-label')
                    .markup('>')
                    .markup('XLSX')
                    .markup('</span>')
                    .markup('</a>')
                    .markup('</li>');
                
                if(!p_region_static_id || current_region_id.match(re) != null)
                 {            
                  $('.a-IRR-iconList').append(html.toString());
                  $dialog_window.find("li.a-IRR-iconList-item a[id$='_XLS']").parent('li').hide();
                 };
                $dialog_window.parent().find("button").addClass("gpvCloseButton");
                
             } else {
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
                $dialog_window.parent().find("button").addClass("gpvCloseButton");
             }     //5.0  

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