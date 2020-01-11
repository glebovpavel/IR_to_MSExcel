CREATE OR REPLACE PACKAGE ir_to_msexcel 
  AUTHID current_user
AS
  PLUGIN_VERSION CONSTANT VARCHAR2(10) DEFAULT '3.20'; 
 
  FUNCTION render  (p_dynamic_action IN apex_plugin.t_dynamic_action,
                    p_plugin         IN apex_plugin.t_plugin )
  RETURN apex_plugin.t_dynamic_action_render_result; 
  
  FUNCTION ajax (p_dynamic_action IN apex_plugin.t_dynamic_action,
                 p_plugin         IN apex_plugin.t_plugin )
  RETURN apex_plugin.t_dynamic_action_ajax_result;

  FUNCTION is_ir2msexcel 
  RETURN BOOLEAN;
  
end IR_TO_MSEXCEL;
/

