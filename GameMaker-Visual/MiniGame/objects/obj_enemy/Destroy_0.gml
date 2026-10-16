/// @DnDAction : YoYo Games.Common.If_Variable
/// @DnDVersion : 1
/// @DnDHash : 7FCFDC2B
/// @DnDApplyTo : {obj_GameManager}
/// @DnDArgument : "var" "enemies_remaining"
/// @DnDArgument : "op" "3"
with(obj_GameManager) var l7FCFDC2B_0 = enemies_remaining <= 0;
if(l7FCFDC2B_0){	/// @DnDAction : YoYo Games.Common.If_Variable
	/// @DnDVersion : 1
	/// @DnDHash : 699D4B86
	/// @DnDApplyTo : {obj_GameManager}
	/// @DnDParent : 7FCFDC2B
	/// @DnDArgument : "var" "wave"
	/// @DnDArgument : "op" "4"
	/// @DnDArgument : "value" "total_waves"
	with(obj_GameManager) var l699D4B86_0 = wave >= total_waves;
	if(l699D4B86_0){	/// @DnDAction : YoYo Games.Common.Variable
		/// @DnDVersion : 1
		/// @DnDHash : 78C2323D
		/// @DnDInput : 2
		/// @DnDApplyTo : {obj_GameManager}
		/// @DnDParent : 699D4B86
		/// @DnDArgument : "expr" "true"
		/// @DnDArgument : "expr_1" "false"
		/// @DnDArgument : "var" "has_won"
		/// @DnDArgument : "var_1" "game_active"
		with(obj_GameManager) {
		has_won = true;
		game_active = false;
		
		}
	
		/// @DnDAction : YoYo Games.Audio.Play_Audio
		/// @DnDVersion : 1.1
		/// @DnDHash : 2791B59F
		/// @DnDParent : 699D4B86
		/// @DnDArgument : "soundid" "sfx_win"
		/// @DnDSaveInfo : "soundid" "sfx_win"
		audio_play_sound(sfx_win, 0, 0, 1.0, undefined, 1.0);}

	/// @DnDAction : YoYo Games.Common.Function_Call
	/// @DnDVersion : 1
	/// @DnDHash : 44C2E05E
	/// @DnDParent : 7FCFDC2B
	/// @DnDArgument : "function" "wave_cleared"
	wave_cleared();}