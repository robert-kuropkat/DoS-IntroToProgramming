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
	/// @DnDApplyTo : {obj_GameManager}
	/// @DnDParent : 3E399392
	/// @DnDArgument : "value" "-1"
	/// @DnDArgument : "value_relative" "1"
	/// @DnDArgument : "instvar" "22"
	with(obj_GameManager) {
	__dnd_lives += -1;
	}

	/// @DnDAction : YoYo Games.Instance Variables.If_Lives
	/// @DnDVersion : 1
	/// @DnDHash : 7EC42A59
	/// @DnDApplyTo : {obj_GameManager}
	/// @DnDParent : 3E399392
	/// @DnDArgument : "op" "3"
	with(obj_GameManager) {
	if(!variable_instance_exists(id, "__dnd_lives")) __dnd_lives = 0;
	var l7EC42A59_0 = __dnd_lives <= 0;
	}
	if(l7EC42A59_0){	/// @DnDAction : YoYo Games.Common.Variable
		/// @DnDVersion : 1
		/// @DnDHash : 0F1F944A
		/// @DnDApplyTo : {obj_GameManager}
		/// @DnDParent : 7EC42A59
		/// @DnDArgument : "expr" "true"
		/// @DnDArgument : "var" "has_lost"
		with(obj_GameManager) {
		has_lost = true;
		
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
	else{	/// @DnDAction : YoYo Games.Audio.Play_Audio
		/// @DnDVersion : 1.1
		/// @DnDHash : 5648389C
		/// @DnDParent : 6509A2BA
		/// @DnDArgument : "soundid" "sfx_lifelost"
		/// @DnDSaveInfo : "soundid" "sfx_lifelost"
		audio_play_sound(sfx_lifelost, 0, 0, 1.0, undefined, 1.0);}

	/// @DnDAction : YoYo Games.Miscellaneous.Debug_Show_Message
	/// @DnDVersion : 1
	/// @DnDHash : 39382721
	/// @DnDParent : 3E399392
	/// @DnDArgument : "msg" ""Enemy Escaped!""
	show_debug_message(string("Enemy Escaped!"));

	/// @DnDAction : YoYo Games.Instances.Destroy_Instance
	/// @DnDVersion : 1
	/// @DnDHash : 5E31200C
	/// @DnDParent : 3E399392
	instance_destroy();}