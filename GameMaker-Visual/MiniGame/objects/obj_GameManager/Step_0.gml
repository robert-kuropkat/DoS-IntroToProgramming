/// @DnDAction : YoYo Games.Common.If_Variable
/// @DnDVersion : 1
/// @DnDHash : 39B4DCE1
/// @DnDApplyTo : {obj_gamemanager}
/// @DnDArgument : "var" "game_active"
/// @DnDArgument : "value" "false"
with(obj_gamemanager) var l39B4DCE1_0 = game_active == false;
if(l39B4DCE1_0){	/// @DnDAction : YoYo Games.Mouse & Keyboard.If_Key_Pressed
	/// @DnDVersion : 1
	/// @DnDHash : 2E2FC178
	/// @DnDParent : 39B4DCE1
	/// @DnDArgument : "key" "ord("R")"
	var l2E2FC178_0;l2E2FC178_0 = keyboard_check_pressed(ord("R"));if (l2E2FC178_0){	/// @DnDAction : YoYo Games.Common.Function_Call
		/// @DnDVersion : 1
		/// @DnDHash : 47B82D62
		/// @DnDParent : 2E2FC178
		/// @DnDArgument : "function" "start_game"
		start_game();}}

/// @DnDAction : YoYo Games.Common.If_Variable
/// @DnDVersion : 1
/// @DnDHash : 3C9337DF
/// @DnDApplyTo : {obj_gamemanager}
/// @DnDArgument : "var" "wave_message_timer"
/// @DnDArgument : "op" "2"
with(obj_gamemanager) var l3C9337DF_0 = wave_message_timer > 0;
if(l3C9337DF_0){	/// @DnDAction : YoYo Games.Common.Variable
	/// @DnDVersion : 1
	/// @DnDHash : 23D607EE
	/// @DnDApplyTo : {obj_gamemanager}
	/// @DnDParent : 3C9337DF
	/// @DnDArgument : "expr" "-1"
	/// @DnDArgument : "expr_relative" "1"
	/// @DnDArgument : "var" "wave_message_timer"
	with(obj_gamemanager) {
	wave_message_timer += -1;
	
	}}