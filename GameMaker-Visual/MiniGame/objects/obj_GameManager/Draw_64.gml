/// @DnDAction : YoYo Games.Drawing.Draw_Value
/// @DnDVersion : 1
/// @DnDHash : 17AABBFD
/// @DnDArgument : "x" "100"
/// @DnDArgument : "y" "50"
/// @DnDArgument : "caption" ""You win!  Enemies remaining: ""
/// @DnDArgument : "var" "enemies_remaining"
draw_text(100, 50, string("You win!  Enemies remaining: ") + string(enemies_remaining));

/// @DnDAction : YoYo Games.Drawing.Draw_Instance_Lives
/// @DnDVersion : 1
/// @DnDHash : 14AAA902
/// @DnDArgument : "x" "100"
/// @DnDArgument : "y" "90"
/// @DnDArgument : "sprite" "Sprite_Blue"
/// @DnDSaveInfo : "sprite" "Sprite_Blue"
var l14AAA902_0 = sprite_get_width(Sprite_Blue);var l14AAA902_1 = 0;if(!variable_instance_exists(id, "__dnd_lives")) __dnd_lives = 0;for(var l14AAA902_2 = __dnd_lives; l14AAA902_2 > 0; --l14AAA902_2) {	draw_sprite(Sprite_Blue, 0, 100 + l14AAA902_1, 90);	l14AAA902_1 += l14AAA902_0;}

/// @DnDAction : YoYo Games.Drawing.Draw_Value
/// @DnDVersion : 1
/// @DnDHash : 3B91082D
/// @DnDArgument : "x" "100"
/// @DnDArgument : "y" "70"
/// @DnDArgument : "caption" ""Player Lives remaining: ""
/// @DnDArgument : "var" "__dnd_lives"
draw_text(100, 70, string("Player Lives remaining: ") + string(__dnd_lives));

/// @DnDAction : YoYo Games.Common.If_Expression
/// @DnDVersion : 1
/// @DnDHash : 0652D2C9
/// @DnDArgument : "expr" "has_won"
if(has_won){	/// @DnDAction : YoYo Games.Drawing.Draw_Value
	/// @DnDVersion : 1
	/// @DnDHash : 3955854D
	/// @DnDParent : 0652D2C9
	/// @DnDArgument : "x" "400"
	/// @DnDArgument : "y" "50"
	/// @DnDArgument : "caption" ""You win!  Enemies remaining: ""
	/// @DnDArgument : "var" "enemies_remaining"
	draw_text(400, 50, string("You win!  Enemies remaining: ") + string(enemies_remaining));}

/// @DnDAction : YoYo Games.Common.If_Expression
/// @DnDVersion : 1
/// @DnDHash : 6062FD39
/// @DnDArgument : "expr" "has_lost"
if(has_lost){	/// @DnDAction : YoYo Games.Drawing.Draw_Value
	/// @DnDVersion : 1
	/// @DnDHash : 0C7E77E1
	/// @DnDParent : 6062FD39
	/// @DnDArgument : "x" "400"
	/// @DnDArgument : "y" "50"
	/// @DnDArgument : "caption" ""You Lose!  Enemies remaining: ""
	/// @DnDArgument : "var" "enemies_remaining"
	draw_text(400, 50, string("You Lose!  Enemies remaining: ") + string(enemies_remaining));}