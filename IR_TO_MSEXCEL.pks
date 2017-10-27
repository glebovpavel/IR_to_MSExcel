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

