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
        if (!apex.region(p_region_static_id)) { /* in case of conditional region is not rendered*/
          return;
        };
        $('body').on( "dialogopen", function( event, ui ) {
          try {
            var $dialog_window = $(event.target);
            //var $dialog_instance = $dialog_window.dialog("instance");   // do not work in APEX 5.2         
            var html;
            
            if (event.target.id !== (p_region_static_id + '_dialog_js')) {
              return;
            }            
            //if( $dialog_instance.options.title !== apex.lang.getMessage( "APEXIR_DOWNLOAD")) {
              if($dialog_window.parent().find('span.ui-dialog-title').text() !== apex.lang.getMessage( "APEXIR_DOWNLOAD"))  {
              return;
            }  
            if(p_version < 201) {
              html = apex.util.htmlBuilder();
              html.markup('<li')
                  .attr('class', 'a-IRR-iconList-item')
                  .markup('>')                    
                  .markup('<a')
                  .attr('class', 'a-IRR-iconList-link')
                  .attr('href','javascript:excel_gpv.getExcel(' + "'" + p_region_static_id + "','" + plugin_id_in +  "'" +');')
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
              $('.a-IRR-iconList').append(html.toString());
              $dialog_window.find("li.a-IRR-iconList-item a[id$='_XLS']").parent('li').hide();
              $dialog_window.parent().find("button").addClass("gpvCloseButton");
            } else {
              /* APEX 20.2 */
              html = apex.util.htmlBuilder();
              html.markup('<button')
                  .attr('type', 'button')
                  .attr('class', 'ui-button--hot ui-button ui-corner-all ui-widget')
                  .attr('onclick','javascript:excel_gpv.getExcel(' + "'" + p_region_static_id + "','" + plugin_id_in +  "'" +');')
                  .attr('style', '  display: block;position: absolute;left: 4px;')
                  .markup('>')    
                  .markup('XLSX')
                  .markup('</button>');

              $(".ui-dialog-buttonset").prepend(html.toString());
            }
          } catch(err) {
            console.info("Error: ",err)
          };    
        });
      }
      
      return {
        getExcel:getExcel
      , addDownloadXLSXIcon: addDownloadXLSXIcon
      };
  
		}();
	}
})(window, apex.jQuery)