/// @DnDAction : YoYo Games.Audio.Play_Audio
/// @DnDVersion : 1.1
/// @DnDHash : 381A859C
/// @DnDArgument : "soundid" "snd_shoot"
/// @DnDSaveInfo : "soundid" "snd_shoot"
audio_play_sound(snd_shoot, 0, 0, 1.0, undefined, 1.0);

/// @DnDAction : YoYo Games.Common.Variable
/// @DnDVersion : 1
/// @DnDHash : 116A4532
/// @DnDArgument : "expr" "6"
/// @DnDArgument : "var" "just_fired"
just_fired = 6;

/// @DnDAction : YoYo Games.Instances.Color_Sprite
/// @DnDVersion : 1
/// @DnDHash : 64B201AA
/// @DnDArgument : "colour" "$FF2692FF"
image_blend = $FF2692FF & $ffffff;
image_alpha = ($FF2692FF >> 24) / $ff;

/// @DnDAction : YoYo Games.Common.Variable
/// @DnDVersion : 1
/// @DnDHash : 777DB68A
/// @DnDInput : 2
/// @DnDArgument : "expr" "1.3"
/// @DnDArgument : "expr_1" "1.3"
/// @DnDArgument : "var" "image_xscale"
/// @DnDArgument : "var_1" "image_yscale"
image_xscale = 1.3;
image_yscale = 1.3;