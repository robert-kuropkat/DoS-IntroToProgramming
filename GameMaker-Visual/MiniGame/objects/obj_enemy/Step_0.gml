/// @DnDAction : YoYo Games.Common.If_Variable
/// @DnDVersion : 1
/// @DnDHash : 5FC9AA3E
/// @DnDApplyTo : {obj_gamemanager}
/// @DnDArgument : "var" "game_active"
/// @DnDArgument : "value" "false"
with(obj_gamemanager) var l5FC9AA3E_0 = game_active == false;
if(l5FC9AA3E_0){	/// @DnDAction : YoYo Games.Common.Exit_Event
	/// @DnDVersion : 1
	/// @DnDHash : 1D7A10D2
	/// @DnDParent : 5FC9AA3E
	exit;}

/// @DnDAction : YoYo Games.Movement.Jump_To_Point
/// @DnDVersion : 1
/// @DnDHash : 1AC268A0
/// @DnDArgument : "x" "0"
/// @DnDArgument : "x_relative" "1"
/// @DnDArgument : "y" "enemy_speed"
/// @DnDArgument : "y_relative" "1"
x += 0;y += enemy_speed;

/// @DnDAction : YoYo Games.Common.If_Expression
/// @DnDVersion : 1
/// @DnDHash : 3E399392
/// @DnDArgument : "expr" "y > room_height"
if(y > room_height){	/// @DnDAction : YoYo Games.Instances.Set_Instance_Var
	/// @DnDVersion : 1
	/// @DnDHash : 64DE6D5E
	/// @DnDApplyTo : {obj_gamemanager}
	/// @DnDParent : 3E399392
	/// @DnDArgument : "value" "-1"
	/// @DnDArgument : "value_relative" "1"
	/// @DnDArgument : "instvar" "22"
	with(obj_gamemanager) {
	__dnd_lives += -1;
	}

	/// @DnDAction : YoYo Games.Common.Variable
	/// @DnDVersion : 1
	/// @DnDHash : 2192FF5F
	/// @DnDApplyTo : {obj_gamemanager}
	/// @DnDParent : 3E399392
	/// @DnDArgument : "expr" "-1"
	/// @DnDArgument : "expr_relative" "1"
	/// @DnDArgument : "var" "enemies_remaining"
	with(obj_gamemanager) {
	enemies_remaining += -1;
	
	}

	/// @DnDAction : YoYo Games.Instance Variables.If_Lives
	/// @DnDVersion : 1
	/// @DnDHash : 7EC42A59
	/// @DnDApplyTo : {obj_gamemanager}
	/// @DnDParent : 3E399392
	/// @DnDArgument : "op" "3"
	with(obj_gamemanager) {
	if(!variable_instance_exists(id, "__dnd_lives")) __dnd_lives = 0;
	var l7EC42A59_0 = __dnd_lives <= 0;
	}
	if(l7EC42A59_0){	/// @DnDAction : YoYo Games.Common.Variable
		/// @DnDVersion : 1
		/// @DnDHash : 0F1F944A
		/// @DnDInput : 2
		/// @DnDApplyTo : {obj_gamemanager}
		/// @DnDParent : 7EC42A59
		/// @DnDArgument : "expr" "true"
		/// @DnDArgument : "expr_1" "false"
		/// @DnDArgument : "var" "has_lost"
		/// @DnDArgument : "var_1" "game_active"
		with(obj_gamemanager) {
		has_lost = true;
		game_active = false;
		
		}
	
		/// @DnDAction : YoYo Games.Audio.Play_Audio
		/// @DnDVersion : 1.1
		/// @DnDHash : 5995F4B1
		/// @DnDParent : 7EC42A59
		/// @DnDArgument : "soundid" "sfx_gameover"
		/// @DnDSaveInfo : "soundid" "sfx_gameover"
		audio_play_sound(sfx_gameover, 0, 0, 1.0, undefined, 1.0);}

	/// @DnDAction : YoYo Games.Common.Else
	/// @DnDVersion : 1
	/// @DnDHash : 6509A2BA
	/// @DnDParent : 3E399392
	else{	/// @DnDAction : YoYo Games.Common.If_Variable
		/// @DnDVersion : 1
		/// @DnDHash : 195CCC7E
		/// @DnDApplyTo : {obj_gamemanager}
		/// @DnDParent : 6509A2BA
		/// @DnDArgument : "var" "enemies_remaining"
		/// @DnDArgument : "op" "3"
		with(obj_gamemanager) var l195CCC7E_0 = enemies_remaining <= 0;
		if(l195CCC7E_0){	/// @DnDAction : YoYo Games.Common.Function_Call
			/// @DnDVersion : 1
			/// @DnDHash : 2ACB8DBB
			/// @DnDParent : 195CCC7E
			/// @DnDArgument : "function" "wave_cleared"
			wave_cleared();}
	
		/// @DnDAction : YoYo Games.Common.Else
		/// @DnDVersion : 1
		/// @DnDHash : 3D70B0C5
		/// @DnDParent : 6509A2BA
		else{	/// @DnDAction : YoYo Games.Audio.Play_Audio
			/// @DnDVersion : 1.1
			/// @DnDHash : 5648389C
			/// @DnDParent : 3D70B0C5
			/// @DnDArgument : "soundid" "sfx_lifelost"
			/// @DnDSaveInfo : "soundid" "sfx_lifelost"
			audio_play_sound(sfx_lifelost, 0, 0, 1.0, undefined, 1.0);}}

	/// @DnDAction : YoYo Games.Instances.Destroy_Instance
	/// @DnDVersion : 1
	/// @DnDHash : 5E31200C
	/// @DnDParent : 3E399392
	instance_destroy();}