/// @DnDAction : YoYo Games.Common.Variable
/// @DnDVersion : 1
/// @DnDHash : 632A67A9
/// @DnDInput : 3
/// @DnDArgument : "expr_1" "false"
/// @DnDArgument : "expr_2" "false"
/// @DnDArgument : "var" "enemies_remaining"
/// @DnDArgument : "var_1" "has_won"
/// @DnDArgument : "var_2" "has_lost"
enemies_remaining = 0;
has_won = false;
has_lost = false;

/// @DnDAction : YoYo Games.Instance Variables.Set_Lives
/// @DnDVersion : 1
/// @DnDHash : 6B13F77B
/// @DnDArgument : "lives" "3"
__dnd_lives = real(3);

/// @DnDAction : YoYo Games.Audio.Play_Audio
/// @DnDVersion : 1.1
/// @DnDHash : 0210DB83
/// @DnDArgument : "soundid" "sfx_start"
/// @DnDSaveInfo : "soundid" "sfx_start"
audio_play_sound(sfx_start, 0, 0, 1.0, undefined, 1.0);