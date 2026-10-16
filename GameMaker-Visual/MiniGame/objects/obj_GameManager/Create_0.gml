/// @DnDAction : YoYo Games.Common.Variable
/// @DnDVersion : 1
/// @DnDHash : 632A67A9
/// @DnDInput : 7
/// @DnDArgument : "expr" "true"
/// @DnDArgument : "expr_1" "false"
/// @DnDArgument : "expr_2" "false"
/// @DnDArgument : "expr_4" "1"
/// @DnDArgument : "expr_5" "3"
/// @DnDArgument : "var" "game_active"
/// @DnDArgument : "var_1" "has_won"
/// @DnDArgument : "var_2" "has_lost"
/// @DnDArgument : "var_3" "enemies_remaining"
/// @DnDArgument : "var_4" "wave"
/// @DnDArgument : "var_5" "total_waves"
/// @DnDArgument : "var_6" "wave_message_timer"
game_active = true;
has_won = false;
has_lost = false;
enemies_remaining = 0;
wave = 1;
total_waves = 3;
wave_message_timer = 0;

/// @DnDAction : YoYo Games.Instance Variables.Set_Lives
/// @DnDVersion : 1
/// @DnDHash : 6B13F77B
/// @DnDArgument : "lives" "3"
__dnd_lives = real(3);

/// @DnDAction : YoYo Games.Instance Variables.Set_Score
/// @DnDVersion : 1
/// @DnDHash : 6D574B59
__dnd_score = real(0);

/// @DnDAction : YoYo Games.Audio.Play_Audio
/// @DnDVersion : 1.1
/// @DnDHash : 0210DB83
/// @DnDArgument : "soundid" "sfx_start"
/// @DnDSaveInfo : "soundid" "sfx_start"
audio_play_sound(sfx_start, 0, 0, 1.0, undefined, 1.0);