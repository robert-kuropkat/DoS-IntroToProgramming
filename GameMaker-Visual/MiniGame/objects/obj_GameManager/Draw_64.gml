/// @DnDAction : YoYo Games.Drawing.Set_Color
/// @DnDVersion : 1
/// @DnDHash : 4D74F391
/// @DnDArgument : "color" "$FF00FFFF"
draw_set_colour($FF00FFFF & $ffffff);
var l4D74F391_0=($FF00FFFF >> 24);
draw_set_alpha(l4D74F391_0 / $ff);

/// @DnDAction : YoYo Games.Drawing.Draw_Value
/// @DnDVersion : 1
/// @DnDHash : 3B91082D
/// @DnDArgument : "x" "20"
/// @DnDArgument : "y" "30"
/// @DnDArgument : "caption" ""Lives: ""
/// @DnDArgument : "var" "__dnd_lives"
draw_text(20, 30, string("Lives: ") + string(__dnd_lives));

/// @DnDAction : YoYo Games.Drawing.Draw_Instance_Score
/// @DnDVersion : 1
/// @DnDHash : 29124F15
/// @DnDArgument : "x" "20"
/// @DnDArgument : "y" "50"
if(!variable_instance_exists(id, "__dnd_score")) __dnd_score = 0;draw_text(20, 50, string("Score: ") + string(__dnd_score));

/// @DnDAction : YoYo Games.Drawing.Draw_Value
/// @DnDVersion : 1
/// @DnDHash : 7487AB3D
/// @DnDApplyTo : {obj_gamemanager}
/// @DnDArgument : "x" "20"
/// @DnDArgument : "y" "70"
/// @DnDArgument : "caption" ""Enemy: ""
/// @DnDArgument : "var" "enemies_remaining"
with(obj_gamemanager) draw_text(20, 70, string("Enemy: ") + string(enemies_remaining));

/// @DnDAction : YoYo Games.Drawing.Draw_Value
/// @DnDVersion : 1
/// @DnDHash : 3572801A
/// @DnDApplyTo : {obj_gamemanager}
/// @DnDArgument : "x" "20"
/// @DnDArgument : "y" "90"
/// @DnDArgument : "caption" ""Wave : " + string(wave) + " of " + string(total_waves)"
with(obj_gamemanager) draw_text(20, 90, string("Wave : " + string(wave) + " of " + string(total_waves)) + "");

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
if(has_won){	/// @DnDAction : YoYo Games.Drawing.Set_Color
	/// @DnDVersion : 1
	/// @DnDHash : 781DAFD1
	/// @DnDParent : 0652D2C9
	/// @DnDArgument : "color" "$FF00FFFF"
	draw_set_colour($FF00FFFF & $ffffff);
	var l781DAFD1_0=($FF00FFFF >> 24);
	draw_set_alpha(l781DAFD1_0 / $ff);

	/// @DnDAction : YoYo Games.Drawing.Draw_Value
	/// @DnDVersion : 1
	/// @DnDHash : 3955854D
	/// @DnDParent : 0652D2C9
	/// @DnDArgument : "x" "400"
	/// @DnDArgument : "y" "50"
	/// @DnDArgument : "caption" ""You win!  Press 'R' to restart.""
	draw_text(400, 50, string("You win!  Press 'R' to restart.") + "");

	/// @DnDAction : YoYo Games.Drawing.Set_Color
	/// @DnDVersion : 1
	/// @DnDHash : 4A07F6B2
	/// @DnDParent : 0652D2C9
	draw_set_colour($FFFFFFFF & $ffffff);
	var l4A07F6B2_0=($FFFFFFFF >> 24);
	draw_set_alpha(l4A07F6B2_0 / $ff);}

/// @DnDAction : YoYo Games.Common.If_Expression
/// @DnDVersion : 1
/// @DnDHash : 6062FD39
/// @DnDArgument : "expr" "has_lost"
if(has_lost){	/// @DnDAction : YoYo Games.Drawing.Set_Color
	/// @DnDVersion : 1
	/// @DnDHash : 3FCF5CAC
	/// @DnDParent : 6062FD39
	/// @DnDArgument : "color" "$FF00FFFF"
	draw_set_colour($FF00FFFF & $ffffff);
	var l3FCF5CAC_0=($FF00FFFF >> 24);
	draw_set_alpha(l3FCF5CAC_0 / $ff);

	/// @DnDAction : YoYo Games.Drawing.Draw_Value
	/// @DnDVersion : 1
	/// @DnDHash : 0C7E77E1
	/// @DnDParent : 6062FD39
	/// @DnDArgument : "x" "400"
	/// @DnDArgument : "y" "50"
	/// @DnDArgument : "caption" ""You Lose!  Press 'R' to restart.""
	draw_text(400, 50, string("You Lose!  Press 'R' to restart.") + "");

	/// @DnDAction : YoYo Games.Drawing.Set_Color
	/// @DnDVersion : 1
	/// @DnDHash : 4D5D01DD
	/// @DnDParent : 6062FD39
	draw_set_colour($FFFFFFFF & $ffffff);
	var l4D5D01DD_0=($FFFFFFFF >> 24);
	draw_set_alpha(l4D5D01DD_0 / $ff);}

/// @DnDAction : YoYo Games.Common.If_Variable
/// @DnDVersion : 1
/// @DnDHash : 009A4CFA
/// @DnDApplyTo : {obj_gamemanager}
/// @DnDArgument : "var" "wave_message_timer"
/// @DnDArgument : "op" "2"
with(obj_gamemanager) var l009A4CFA_0 = wave_message_timer > 0;
if(l009A4CFA_0){	/// @DnDAction : YoYo Games.Drawing.Set_Color
	/// @DnDVersion : 1
	/// @DnDHash : 00DEAA95
	/// @DnDParent : 009A4CFA
	/// @DnDArgument : "color" "$FF00FFFF"
	draw_set_colour($FF00FFFF & $ffffff);
	var l00DEAA95_0=($FF00FFFF >> 24);
	draw_set_alpha(l00DEAA95_0 / $ff);

	/// @DnDAction : YoYo Games.Drawing.Draw_Value
	/// @DnDVersion : 1
	/// @DnDHash : 03728171
	/// @DnDApplyTo : {obj_gamemanager}
	/// @DnDParent : 009A4CFA
	/// @DnDArgument : "x" "400"
	/// @DnDArgument : "y" "50"
	/// @DnDArgument : "caption" ""Wave ""
	/// @DnDArgument : "var" "wave"
	with(obj_gamemanager) draw_text(400, 50, string("Wave ") + string(wave));

	/// @DnDAction : YoYo Games.Drawing.Set_Color
	/// @DnDVersion : 1
	/// @DnDHash : 4BCB4BDD
	/// @DnDParent : 009A4CFA
	draw_set_colour($FFFFFFFF & $ffffff);
	var l4BCB4BDD_0=($FFFFFFFF >> 24);
	draw_set_alpha(l4BCB4BDD_0 / $ff);}