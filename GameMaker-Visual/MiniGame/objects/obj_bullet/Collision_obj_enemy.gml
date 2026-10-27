/// @DnDAction : YoYo Games.Audio.Play_Audio
/// @DnDVersion : 1.1
/// @DnDHash : 71B4082D
/// @DnDArgument : "soundid" "snd_enemyHit"
/// @DnDSaveInfo : "soundid" "snd_enemyHit"
audio_play_sound(snd_enemyHit, 0, 0, 1.0, undefined, 1.0);

/// @DnDAction : YoYo Games.Instances.Create_Instance
/// @DnDVersion : 1
/// @DnDHash : 4EE2730C
/// @DnDApplyTo : other
/// @DnDArgument : "xpos" "x"
/// @DnDArgument : "ypos" "y"
/// @DnDArgument : "objectid" "obj_explosion"
/// @DnDSaveInfo : "objectid" "obj_explosion"
with(other) {
	instance_create_layer(x, y, "Instances", obj_explosion); 
}

/// @DnDAction : YoYo Games.Instances.Set_Sprite
/// @DnDVersion : 1
/// @DnDHash : 1D2FDC24
/// @DnDApplyTo : other
/// @DnDArgument : "spriteind" "spr_enemy_flash"
/// @DnDSaveInfo : "spriteind" "spr_enemy_flash"
with(other) {
sprite_index = spr_enemy_flash;
image_index = 0;
}

/// @DnDAction : YoYo Games.Instances.Color_Sprite
/// @DnDVersion : 1
/// @DnDHash : 00F2302B
/// @DnDApplyTo : other
/// @DnDArgument : "colour" "c_white"
/// @DnDArgument : "alpha" "false"
with(other) {
image_blend = c_white & $ffffff;
}

/// @DnDAction : YoYo Games.Instances.Set_Alarm
/// @DnDVersion : 1
/// @DnDHash : 0C65697B
/// @DnDApplyTo : other
/// @DnDArgument : "steps" "6"
with(other) {
alarm_set(0, 6);

}

/// @DnDAction : YoYo Games.Common.Variable
/// @DnDVersion : 1
/// @DnDHash : 41D0DDE1
/// @DnDApplyTo : {obj_gamemanager}
/// @DnDArgument : "expr" "-1"
/// @DnDArgument : "expr_relative" "1"
/// @DnDArgument : "var" "enemies_remaining"
with(obj_gamemanager) {
enemies_remaining += -1;

}

/// @DnDAction : YoYo Games.Instance Variables.Set_Score
/// @DnDVersion : 1
/// @DnDHash : 61E91593
/// @DnDApplyTo : {obj_gamemanager}
/// @DnDArgument : "score" "10"
/// @DnDArgument : "score_relative" "1"
with(obj_gamemanager) {
if(!variable_instance_exists(id, "__dnd_score")) __dnd_score = 0;__dnd_score += real(10);
}

/// @DnDAction : YoYo Games.Common.If_Variable
/// @DnDVersion : 1
/// @DnDHash : 4D19CB85
/// @DnDApplyTo : {obj_gamemanager}
/// @DnDArgument : "var" "enemies_remaining"
/// @DnDArgument : "op" "3"
with(obj_gamemanager) var l4D19CB85_0 = enemies_remaining <= 0;
if(l4D19CB85_0){	/// @DnDAction : YoYo Games.Common.Function_Call
	/// @DnDVersion : 1
	/// @DnDHash : 4D82BE32
	/// @DnDParent : 4D19CB85
	/// @DnDArgument : "function" "wave_cleared"
	wave_cleared();}

/// @DnDAction : YoYo Games.Instances.Destroy_Instance
/// @DnDVersion : 1
/// @DnDHash : 3394871C
instance_destroy();