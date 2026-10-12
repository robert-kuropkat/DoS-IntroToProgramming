/// @DnDAction : YoYo Games.Audio.Play_Audio
/// @DnDVersion : 1.1
/// @DnDHash : 705F45FF
/// @DnDArgument : "soundid" "snd_enemyHit"
/// @DnDSaveInfo : "soundid" "snd_enemyHit"
audio_play_sound(snd_enemyHit, 0, 0, 1.0, undefined, 1.0);

/// @DnDAction : YoYo Games.Common.Variable
/// @DnDVersion : 1
/// @DnDHash : 7B09E65B
/// @DnDArgument : "expr" "-1"
/// @DnDArgument : "expr_relative" "1"
/// @DnDArgument : "var" "obj_GameManager.enemies_remaining"
obj_GameManager.enemies_remaining += -1;

/// @DnDAction : YoYo Games.Instance Variables.Set_Score
/// @DnDVersion : 1
/// @DnDHash : 6D15CB81
/// @DnDApplyTo : {obj_GameManager}
/// @DnDArgument : "score" "10"
/// @DnDArgument : "score_relative" "1"
with(obj_GameManager) {
if(!variable_instance_exists(id, "__dnd_score")) __dnd_score = 0;__dnd_score += real(10);
}

/// @DnDAction : YoYo Games.Instances.Destroy_Instance
/// @DnDVersion : 1
/// @DnDHash : 581C5DFD
instance_destroy();