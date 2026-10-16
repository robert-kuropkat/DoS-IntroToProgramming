/// @DnDAction : YoYo Games.Drawing.Draw_Instance_Lives
/// @DnDVersion : 1
/// @DnDHash : 14AAA902
/// @DnDArgument : "x" "100"
/// @DnDArgument : "y" "90"
/// @DnDArgument : "sprite" "Sprite_Blue"
/// @DnDSaveInfo : "sprite" "Sprite_Blue"
var l14AAA902_0 = sprite_get_width(Sprite_Blue);var l14AAA902_1 = 0;if(!variable_instance_exists(id, "__dnd_lives")) __dnd_lives = 0;for(var l14AAA902_2 = __dnd_lives; l14AAA902_2 > 0; --l14AAA902_2) {	draw_sprite(Sprite_Blue, 0, 100 + l14AAA902_1, 90);	l14AAA902_1 += l14AAA902_0;}

/// @DnDAction : YoYo Games.Drawing.Set_Color
/// @DnDVersion : 1
/// @DnDHash : 4D74F391
/// @DnDArgument : "color" "$FF00FFFF"
draw_set_colour($FF00FFFF & $ffffff);
var l4D74F391_0=($FF00FFFF >> 24);
draw_set_alpha(l4D74F391_0 / $ff);

/// @DnDAction : YoYo Games.Drawing.Draw_Instance_Score
/// @DnDVersion : 1
/// @DnDHash : 29124F15
/// @DnDArgument : "x" "20"
/// @DnDArgument : "y" "50"
if(!variable_instance_exists(id, "__dnd_score")) __dnd_score = 0;draw_text(20, 50, string("Score: ") + string(__dnd_score));

/// @DnDAction : YoYo Games.Drawing.Draw_Value
/// @DnDVersion : 1
/// @DnDHash : 3B91082D
/// @DnDArgument : "x" "20"
/// @DnDArgument : "y" "20"
/// @DnDArgument : "caption" ""Player Lives remaining: ""
/// @DnDArgument : "var" "__dnd_lives"
draw_text(20, 20, string("Player Lives remaining: ") + string(__dnd_lives));

/// @DnDAction : YoYo Games.Drawing.Set_Color
/// @DnDVersion : 1
/// @DnDHash : 4D750A8E
draw_set_colour($FFFFFFFF & $ffffff);
var l4D750A8E_0=($FFFFFFFF >> 24);
draw_set_alpha(l4D750A8E_0 / $ff);

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
	/// @DnDArgument : "caption" ""You win!  Press 'R' to restart.""
	draw_text(400, 50, string("You win!  Press 'R' to restart.") + "");}

/// @DnDAction : YoYo Games.Common.If_Expression
/// @DnDVersion : 1
/// @DnDHash : 6062FD39
/// @DnDArgument : "expr" "has_lost"
if(has_lost){	/// @DnDAction : YoYo Games.Drawing.Draw_Value
	/// @DnDVersion : 1
	/// @DnDHash : 0C7E77E1
	/// @DnDParent : 6062FD39
	/// @DnDArgument : "x" "400"
	/// @DnDArgument : "y" "90"
	/// @DnDArgument : "caption" ""You Lose!  Press 'R' to restart.""
	draw_text(400, 90, string("You Lose!  Press 'R' to restart.") + "");}

/// @DnDAction : YoYo Games.Common.If_Variable
/// @DnDVersion : 1
/// @DnDHash : 009A4CFA
/// @DnDApplyTo : {obj_GameManager}
/// @DnDArgument : "var" "wave_message_timer"
/// @DnDArgument : "op" "2"
with(obj_GameManager) var l009A4CFA_0 = wave_message_timer > 0;
if(l009A4CFA_0){	/// @DnDAction : YoYo Games.Drawing.Draw_Value
	/// @DnDVersion : 1
	/// @DnDHash : 03728171
	/// @DnDApplyTo : {obj_GameManager}
	/// @DnDParent : 009A4CFA
	/// @DnDArgument : "caption" ""wave ""
	/// @DnDArgument : "var" "wave"
	with(obj_GameManager) draw_text(0, 0, string("wave ") + string(wave));}