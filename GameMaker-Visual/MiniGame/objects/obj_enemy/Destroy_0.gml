/// @DnDAction : YoYo Games.Common.If_Variable
/// @DnDVersion : 1
/// @DnDHash : 7FCFDC2B
/// @DnDArgument : "var" "obj_GameManager.enemies_remaining"
/// @DnDArgument : "op" "3"
if(obj_GameManager.enemies_remaining <= 0){	/// @DnDAction : YoYo Games.Common.Variable
	/// @DnDVersion : 1
	/// @DnDHash : 78C2323D
	/// @DnDParent : 7FCFDC2B
	/// @DnDArgument : "expr" "true"
	/// @DnDArgument : "var" "obj_GameManager.has_won"
	obj_GameManager.has_won = true;

	/// @DnDAction : YoYo Games.Audio.Play_Audio
	/// @DnDVersion : 1.1
	/// @DnDHash : 2791B59F
	/// @DnDParent : 7FCFDC2B
	/// @DnDArgument : "soundid" "sfx_win"
	/// @DnDSaveInfo : "soundid" "sfx_win"
	audio_play_sound(sfx_win, 0, 0, 1.0, undefined, 1.0);}