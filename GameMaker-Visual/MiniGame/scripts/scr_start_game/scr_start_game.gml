/// @DnDAction : YoYo Games.Common.Function
/// @DnDVersion : 1
/// @DnDHash : 60377EA3
/// @DnDComment : 
/// @DnDArgument : "funcName" "start_game"
function start_game() {	/// @DnDAction : YoYo Games.Common.Apply_To
	/// @DnDVersion : 1
	/// @DnDHash : 07A24CA2
	/// @DnDApplyTo : {obj_enemy}
	/// @DnDParent : 60377EA3
	with(obj_enemy) {
		/// @DnDAction : YoYo Games.Instances.Destroy_Instance
		/// @DnDVersion : 1
		/// @DnDHash : 4F1A6E17
		/// @DnDApplyTo : {obj_enemy}
		/// @DnDParent : 07A24CA2
		with(obj_enemy) instance_destroy();
	}

	/// @DnDAction : YoYo Games.Instance Variables.Set_Score
	/// @DnDVersion : 1
	/// @DnDHash : 48973A6D
	/// @DnDApplyTo : {obj_gamemanager}
	/// @DnDParent : 60377EA3
	with(obj_gamemanager) {
	__dnd_score = real(0);
	}

	/// @DnDAction : YoYo Games.Instance Variables.Set_Lives
	/// @DnDVersion : 1
	/// @DnDHash : 78006E9C
	/// @DnDApplyTo : {obj_gamemanager}
	/// @DnDParent : 60377EA3
	/// @DnDArgument : "lives" "3"
	with(obj_gamemanager) {
	__dnd_lives = real(3);
	}

	/// @DnDAction : YoYo Games.Common.Variable
	/// @DnDVersion : 1
	/// @DnDHash : 446D1A1F
	/// @DnDInput : 6
	/// @DnDApplyTo : {obj_gamemanager}
	/// @DnDParent : 60377EA3
	/// @DnDArgument : "expr" "1"
	/// @DnDArgument : "expr_1" "false"
	/// @DnDArgument : "expr_2" "false"
	/// @DnDArgument : "expr_3" "true"
	/// @DnDArgument : "expr_5" "4 + wave"
	/// @DnDArgument : "var" "wave"
	/// @DnDArgument : "var_1" "has_won"
	/// @DnDArgument : "var_2" "has_lost"
	/// @DnDArgument : "var_3" "game_active"
	/// @DnDArgument : "var_4" "wave_message_timer"
	/// @DnDArgument : "var_5" "enemies_remaining"
	with(obj_gamemanager) {
	wave = 1;
	has_won = false;
	has_lost = false;
	game_active = true;
	wave_message_timer = 0;
	enemies_remaining = 4 + wave;
	
	}

	/// @DnDAction : YoYo Games.Common.Function_Call
	/// @DnDVersion : 1
	/// @DnDHash : 75026C2F
	/// @DnDParent : 60377EA3
	/// @DnDArgument : "function" "spawn_enemies"
	spawn_enemies();}

/// @DnDAction : YoYo Games.Common.Function
/// @DnDVersion : 1
/// @DnDHash : 2A2F9D51
/// @DnDArgument : "funcName" "wave_cleared"
function wave_cleared() {	/// @DnDAction : YoYo Games.Common.If_Variable
	/// @DnDVersion : 1
	/// @DnDHash : 6DA4D33F
	/// @DnDApplyTo : {obj_gamemanager}
	/// @DnDParent : 2A2F9D51
	/// @DnDArgument : "var" "wave"
	/// @DnDArgument : "op" "4"
	/// @DnDArgument : "value" "total_waves"
	with(obj_gamemanager) var l6DA4D33F_0 = wave >= total_waves;
	if(l6DA4D33F_0){	/// @DnDAction : YoYo Games.Common.Variable
		/// @DnDVersion : 1
		/// @DnDHash : 5ED368C0
		/// @DnDInput : 2
		/// @DnDApplyTo : {obj_gamemanager}
		/// @DnDParent : 6DA4D33F
		/// @DnDArgument : "expr" "true"
		/// @DnDArgument : "expr_1" "false"
		/// @DnDArgument : "var" "has_won"
		/// @DnDArgument : "var_1" "game_active"
		with(obj_gamemanager) {
		has_won = true;
		game_active = false;
		
		}
	
		/// @DnDAction : YoYo Games.Audio.Play_Audio
		/// @DnDVersion : 1.1
		/// @DnDHash : 1945EFCE
		/// @DnDParent : 6DA4D33F
		/// @DnDArgument : "soundid" "sfx_win"
		/// @DnDSaveInfo : "soundid" "sfx_win"
		audio_play_sound(sfx_win, 0, 0, 1.0, undefined, 1.0);}

	/// @DnDAction : YoYo Games.Common.Else
	/// @DnDVersion : 1
	/// @DnDHash : 4420BE72
	/// @DnDParent : 2A2F9D51
	else{	/// @DnDAction : YoYo Games.Common.Variable
		/// @DnDVersion : 1
		/// @DnDHash : 678D967A
		/// @DnDInput : 3
		/// @DnDApplyTo : {obj_gamemanager}
		/// @DnDParent : 4420BE72
		/// @DnDArgument : "expr" "1"
		/// @DnDArgument : "expr_relative" "1"
		/// @DnDArgument : "expr_1" "4 + wave"
		/// @DnDArgument : "expr_2" "game_get_speed(gamespeed_fps)"
		/// @DnDArgument : "var" "wave"
		/// @DnDArgument : "var_1" "enemies_remaining"
		/// @DnDArgument : "var_2" "wave_message_timer"
		with(obj_gamemanager) {
		wave += 1;
		enemies_remaining = 4 + wave;
		wave_message_timer = game_get_speed(gamespeed_fps);
		
		}
	
		/// @DnDAction : YoYo Games.Audio.Play_Audio
		/// @DnDVersion : 1.1
		/// @DnDHash : 48DBA2DB
		/// @DnDParent : 4420BE72
		/// @DnDArgument : "soundid" "sfx_wave"
		/// @DnDSaveInfo : "soundid" "sfx_wave"
		audio_play_sound(sfx_wave, 0, 0, 1.0, undefined, 1.0);
	
		/// @DnDAction : YoYo Games.Common.Function_Call
		/// @DnDVersion : 1
		/// @DnDHash : 5CFADC68
		/// @DnDParent : 4420BE72
		/// @DnDArgument : "function" "spawn_enemies"
		spawn_enemies();}}