/// @DnDAction : YoYo Games.Common.If_Variable
/// @DnDVersion : 1
/// @DnDHash : 39B4DCE1
/// @DnDApplyTo : {obj_GameManager}
/// @DnDArgument : "var" "game_active"
/// @DnDArgument : "value" "false"
with(obj_GameManager) var l39B4DCE1_0 = game_active == false;
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