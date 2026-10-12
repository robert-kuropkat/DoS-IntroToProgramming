/// @DnDAction : YoYo Games.Movement.Jump_To_Point
/// @DnDVersion : 1
/// @DnDHash : 19C449BA
/// @DnDArgument : "x" "0"
/// @DnDArgument : "x_relative" "1"
/// @DnDArgument : "y" "-8"
/// @DnDArgument : "y_relative" "1"
x += 0;y += -8;

/// @DnDAction : YoYo Games.Common.If_Variable
/// @DnDVersion : 1
/// @DnDHash : 4EBD27EF
/// @DnDArgument : "var" "just_fired"
/// @DnDArgument : "op" "3"
if(just_fired <= 0){	/// @DnDAction : YoYo Games.Instances.Color_Sprite
	/// @DnDVersion : 1
	/// @DnDHash : 1C4C738D
	/// @DnDParent : 4EBD27EF
	image_blend = $FFFFFFFF & $ffffff;
	image_alpha = ($FFFFFFFF >> 24) / $ff;

	/// @DnDAction : YoYo Games.Common.Variable
	/// @DnDVersion : 1
	/// @DnDHash : 607F4189
	/// @DnDInput : 2
	/// @DnDParent : 4EBD27EF
	/// @DnDArgument : "expr" "1"
	/// @DnDArgument : "expr_1" "1"
	/// @DnDArgument : "var" "image_xscale"
	/// @DnDArgument : "var_1" "image_yscale"
	image_xscale = 1;
	image_yscale = 1;}

/// @DnDAction : YoYo Games.Common.Else
/// @DnDVersion : 1
/// @DnDHash : 38B602F7
else{	/// @DnDAction : YoYo Games.Common.Variable
	/// @DnDVersion : 1
	/// @DnDHash : 4F7A4EE9
	/// @DnDParent : 38B602F7
	/// @DnDArgument : "expr" "-1"
	/// @DnDArgument : "expr_relative" "1"
	/// @DnDArgument : "var" "just_fired"
	just_fired += -1;}

/// @DnDAction : YoYo Games.Common.If_Variable
/// @DnDVersion : 1
/// @DnDHash : 00CD3A1F
/// @DnDArgument : "var" "y"
/// @DnDArgument : "op" "1"
if(y < 0){	/// @DnDAction : YoYo Games.Instances.Destroy_Instance
	/// @DnDVersion : 1
	/// @DnDHash : 7E4471D0
	/// @DnDParent : 00CD3A1F
	instance_destroy();}