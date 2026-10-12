/// @DnDAction : YoYo Games.Common.Temp_Variable
/// @DnDVersion : 1
/// @DnDHash : 5E789A7B
/// @DnDArgument : "var" "moving"
/// @DnDArgument : "value" "false"
var moving = false;

/// @DnDAction : YoYo Games.Common.Variable
/// @DnDVersion : 1
/// @DnDHash : 3372F0CC
/// @DnDArgument : "expr" "-1"
/// @DnDArgument : "expr_relative" "1"
/// @DnDArgument : "var" "fire_cooldown"
fire_cooldown += -1;

/// @DnDAction : YoYo Games.Mouse & Keyboard.If_Key_Pressed
/// @DnDVersion : 1
/// @DnDHash : 5C21B212
var l5C21B212_0;l5C21B212_0 = keyboard_check_pressed(vk_space);if (l5C21B212_0){	/// @DnDAction : YoYo Games.Common.If_Variable
	/// @DnDVersion : 1
	/// @DnDHash : 029CFAB1
	/// @DnDParent : 5C21B212
	/// @DnDArgument : "var" "fire_cooldown"
	/// @DnDArgument : "op" "3"
	if(fire_cooldown <= 0){	/// @DnDAction : YoYo Games.Instances.Create_Instance
		/// @DnDVersion : 1
		/// @DnDHash : 7584F23F
		/// @DnDParent : 029CFAB1
		/// @DnDArgument : "xpos" "+65"
		/// @DnDArgument : "xpos_relative" "1"
		/// @DnDArgument : "ypos_relative" "1"
		/// @DnDArgument : "objectid" "obj_bullet"
		/// @DnDSaveInfo : "objectid" "obj_bullet"
		instance_create_layer(x + +65, y + 0, "Instances", obj_bullet);
	
		/// @DnDAction : YoYo Games.Common.Variable
		/// @DnDVersion : 1
		/// @DnDHash : 41B0DC0D
		/// @DnDParent : 029CFAB1
		/// @DnDArgument : "expr" "12"
		/// @DnDArgument : "var" "fire_cooldown"
		fire_cooldown = 12;}}

/// @DnDAction : YoYo Games.Mouse & Keyboard.If_Key_Down
/// @DnDVersion : 1
/// @DnDHash : 6E381896
/// @DnDArgument : "key" "vk_right"
var l6E381896_0;l6E381896_0 = keyboard_check(vk_right);if (l6E381896_0){	/// @DnDAction : YoYo Games.Common.Variable
	/// @DnDVersion : 1
	/// @DnDHash : 615C1E0C
	/// @DnDParent : 6E381896
	/// @DnDArgument : "expr" "true"
	/// @DnDArgument : "var" "moving"
	moving = true;

	/// @DnDAction : YoYo Games.Movement.Jump_To_Point
	/// @DnDVersion : 1
	/// @DnDHash : 6F59F777
	/// @DnDParent : 6E381896
	/// @DnDArgument : "x" "4"
	/// @DnDArgument : "x_relative" "1"
	/// @DnDArgument : "y" "0"
	/// @DnDArgument : "y_relative" "1"
	x += 4;y += 0;

	/// @DnDAction : YoYo Games.Instances.Set_Sprite
	/// @DnDVersion : 1
	/// @DnDHash : 1163E04C
	/// @DnDParent : 6E381896
	/// @DnDArgument : "spriteind" "Sprite_Yellow"
	/// @DnDSaveInfo : "spriteind" "Sprite_Yellow"
	sprite_index = Sprite_Yellow;
	image_index = 0;

	/// @DnDAction : YoYo Games.Audio.If_Audio_Playing
	/// @DnDVersion : 1
	/// @DnDHash : 15CFC96C
	/// @DnDParent : 6E381896
	/// @DnDArgument : "soundid" "snd_blip"
	/// @DnDArgument : "not" "1"
	/// @DnDSaveInfo : "soundid" "snd_blip"
	var l15CFC96C_0 = snd_blip;if (!audio_is_playing(l15CFC96C_0)){	/// @DnDAction : YoYo Games.Audio.Play_Audio
		/// @DnDVersion : 1.1
		/// @DnDHash : 2772D55A
		/// @DnDParent : 15CFC96C
		/// @DnDArgument : "soundid" "snd_blip"
		/// @DnDSaveInfo : "soundid" "snd_blip"
		audio_play_sound(snd_blip, 0, 0, 1.0, undefined, 1.0);}}

/// @DnDAction : YoYo Games.Mouse & Keyboard.If_Key_Down
/// @DnDVersion : 1
/// @DnDHash : 69F41297
/// @DnDArgument : "key" "vk_left"
var l69F41297_0;l69F41297_0 = keyboard_check(vk_left);if (l69F41297_0){	/// @DnDAction : YoYo Games.Common.Variable
	/// @DnDVersion : 1
	/// @DnDHash : 5CFF3E5D
	/// @DnDParent : 69F41297
	/// @DnDArgument : "expr" "true"
	/// @DnDArgument : "var" "moving"
	moving = true;

	/// @DnDAction : YoYo Games.Movement.Jump_To_Point
	/// @DnDVersion : 1
	/// @DnDHash : 31BDB53C
	/// @DnDParent : 69F41297
	/// @DnDArgument : "x" "-4"
	/// @DnDArgument : "x_relative" "1"
	/// @DnDArgument : "y" "0"
	/// @DnDArgument : "y_relative" "1"
	x += -4;y += 0;

	/// @DnDAction : YoYo Games.Instances.Set_Sprite
	/// @DnDVersion : 1
	/// @DnDHash : 003E7BD2
	/// @DnDParent : 69F41297
	/// @DnDArgument : "spriteind" "Sprite_Green"
	/// @DnDSaveInfo : "spriteind" "Sprite_Green"
	sprite_index = Sprite_Green;
	image_index = 0;

	/// @DnDAction : YoYo Games.Audio.If_Audio_Playing
	/// @DnDVersion : 1
	/// @DnDHash : 598E0CA6
	/// @DnDParent : 69F41297
	/// @DnDArgument : "soundid" "snd_blip"
	/// @DnDArgument : "not" "1"
	/// @DnDSaveInfo : "soundid" "snd_blip"
	var l598E0CA6_0 = snd_blip;if (!audio_is_playing(l598E0CA6_0)){	/// @DnDAction : YoYo Games.Audio.Play_Audio
		/// @DnDVersion : 1.1
		/// @DnDHash : 10EB8234
		/// @DnDParent : 598E0CA6
		/// @DnDArgument : "soundid" "snd_blip"
		/// @DnDSaveInfo : "soundid" "snd_blip"
		audio_play_sound(snd_blip, 0, 0, 1.0, undefined, 1.0);}}

/// @DnDAction : YoYo Games.Mouse & Keyboard.If_Key_Down
/// @DnDVersion : 1
/// @DnDHash : 01606FC4
/// @DnDArgument : "key" "vk_up"
var l01606FC4_0;l01606FC4_0 = keyboard_check(vk_up);if (l01606FC4_0){	/// @DnDAction : YoYo Games.Common.Variable
	/// @DnDVersion : 1
	/// @DnDHash : 33E22D8E
	/// @DnDParent : 01606FC4
	/// @DnDArgument : "expr" "true"
	/// @DnDArgument : "var" "moving"
	moving = true;

	/// @DnDAction : YoYo Games.Movement.Jump_To_Point
	/// @DnDVersion : 1
	/// @DnDHash : 43AB67A6
	/// @DnDParent : 01606FC4
	/// @DnDArgument : "x" "0"
	/// @DnDArgument : "x_relative" "1"
	/// @DnDArgument : "y" "-4"
	/// @DnDArgument : "y_relative" "1"
	x += 0;y += -4;

	/// @DnDAction : YoYo Games.Instances.Set_Sprite
	/// @DnDVersion : 1
	/// @DnDHash : 36737F10
	/// @DnDParent : 01606FC4
	/// @DnDArgument : "spriteind" "Sprite_Red"
	/// @DnDSaveInfo : "spriteind" "Sprite_Red"
	sprite_index = Sprite_Red;
	image_index = 0;

	/// @DnDAction : YoYo Games.Audio.If_Audio_Playing
	/// @DnDVersion : 1
	/// @DnDHash : 1AE4AB81
	/// @DnDParent : 01606FC4
	/// @DnDArgument : "soundid" "snd_blip"
	/// @DnDArgument : "not" "1"
	/// @DnDSaveInfo : "soundid" "snd_blip"
	var l1AE4AB81_0 = snd_blip;if (!audio_is_playing(l1AE4AB81_0)){	/// @DnDAction : YoYo Games.Audio.Play_Audio
		/// @DnDVersion : 1.1
		/// @DnDHash : 2F65FD87
		/// @DnDParent : 1AE4AB81
		/// @DnDArgument : "soundid" "snd_blip"
		/// @DnDSaveInfo : "soundid" "snd_blip"
		audio_play_sound(snd_blip, 0, 0, 1.0, undefined, 1.0);}}

/// @DnDAction : YoYo Games.Mouse & Keyboard.If_Key_Down
/// @DnDVersion : 1
/// @DnDHash : 20DAA7ED
/// @DnDArgument : "key" "vk_down"
var l20DAA7ED_0;l20DAA7ED_0 = keyboard_check(vk_down);if (l20DAA7ED_0){	/// @DnDAction : YoYo Games.Common.Variable
	/// @DnDVersion : 1
	/// @DnDHash : 4C267CCD
	/// @DnDParent : 20DAA7ED
	/// @DnDArgument : "expr" "true"
	/// @DnDArgument : "var" "moving"
	moving = true;

	/// @DnDAction : YoYo Games.Movement.Jump_To_Point
	/// @DnDVersion : 1
	/// @DnDHash : 74BE3746
	/// @DnDParent : 20DAA7ED
	/// @DnDArgument : "x" "0"
	/// @DnDArgument : "x_relative" "1"
	/// @DnDArgument : "y" "4"
	/// @DnDArgument : "y_relative" "1"
	x += 0;y += 4;

	/// @DnDAction : YoYo Games.Instances.Set_Sprite
	/// @DnDVersion : 1
	/// @DnDHash : 1DED81FD
	/// @DnDParent : 20DAA7ED
	/// @DnDArgument : "spriteind" "Sprite_Orange"
	/// @DnDSaveInfo : "spriteind" "Sprite_Orange"
	sprite_index = Sprite_Orange;
	image_index = 0;

	/// @DnDAction : YoYo Games.Audio.If_Audio_Playing
	/// @DnDVersion : 1
	/// @DnDHash : 14017D97
	/// @DnDParent : 20DAA7ED
	/// @DnDArgument : "soundid" "snd_blip"
	/// @DnDArgument : "not" "1"
	/// @DnDSaveInfo : "soundid" "snd_blip"
	var l14017D97_0 = snd_blip;if (!audio_is_playing(l14017D97_0)){	/// @DnDAction : YoYo Games.Audio.Play_Audio
		/// @DnDVersion : 1.1
		/// @DnDHash : 62E92FDD
		/// @DnDParent : 14017D97
		/// @DnDArgument : "soundid" "snd_blip"
		/// @DnDSaveInfo : "soundid" "snd_blip"
		audio_play_sound(snd_blip, 0, 0, 1.0, undefined, 1.0);}}

/// @DnDAction : YoYo Games.Instances.Set_Instance_Var
/// @DnDVersion : 1
/// @DnDHash : 38CDADE3
/// @DnDInput : 2
/// @DnDArgument : "value" "clamp(x, 0, room_width-130)"
/// @DnDArgument : "value_1" "clamp(y, 0, room_height-130)"
/// @DnDArgument : "instvar_1" "1"
x = clamp(x, 0, room_width-130);
y = clamp(y, 0, room_height-130);

/// @DnDAction : YoYo Games.Common.If_Expression
/// @DnDVersion : 1
/// @DnDHash : 15A3CF18
/// @DnDArgument : "expr" "moving"
if(moving){	/// @DnDAction : YoYo Games.Common.Variable
	/// @DnDVersion : 1
	/// @DnDHash : 12C8580B
	/// @DnDParent : 15A3CF18
	/// @DnDArgument : "expr" "1"
	/// @DnDArgument : "expr_relative" "1"
	/// @DnDArgument : "var" "blink_timer"
	blink_timer += 1;

	/// @DnDAction : YoYo Games.Common.If_Expression
	/// @DnDVersion : 1
	/// @DnDHash : 53C22D92
	/// @DnDParent : 15A3CF18
	/// @DnDArgument : "expr" "blink_timer mod 20 < 10"
	/// @DnDArgument : "not" "1"
	if(!(blink_timer mod 20 < 10)){	/// @DnDAction : YoYo Games.Instances.Set_Instance_Var
		/// @DnDVersion : 1
		/// @DnDHash : 3BD5B829
		/// @DnDParent : 53C22D92
		/// @DnDArgument : "value" "false"
		/// @DnDArgument : "instvar" "6"
		visible = false;}

	/// @DnDAction : YoYo Games.Common.Else
	/// @DnDVersion : 1
	/// @DnDHash : 40030349
	/// @DnDParent : 15A3CF18
	else{	/// @DnDAction : YoYo Games.Instances.Set_Instance_Var
		/// @DnDVersion : 1
		/// @DnDHash : 6F54A3E4
		/// @DnDParent : 40030349
		/// @DnDArgument : "value" "true"
		/// @DnDArgument : "instvar" "6"
		visible = true;}}

/// @DnDAction : YoYo Games.Common.Else
/// @DnDVersion : 1
/// @DnDHash : 5B4CFAD8
else{	/// @DnDAction : YoYo Games.Instances.Set_Sprite
	/// @DnDVersion : 1
	/// @DnDHash : 261F3415
	/// @DnDParent : 5B4CFAD8
	/// @DnDArgument : "spriteind" "Sprite_Blue"
	/// @DnDSaveInfo : "spriteind" "Sprite_Blue"
	sprite_index = Sprite_Blue;
	image_index = 0;

	/// @DnDAction : YoYo Games.Instances.Set_Instance_Var
	/// @DnDVersion : 1
	/// @DnDHash : 65C34F48
	/// @DnDParent : 5B4CFAD8
	/// @DnDArgument : "value" "true"
	/// @DnDArgument : "instvar" "6"
	visible = true;}