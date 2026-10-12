/// @DnDAction : YoYo Games.Common.Temp_Variable
/// @DnDVersion : 1
/// @DnDHash : 66BD70F4
/// @DnDArgument : "var" "enemy_count"
/// @DnDArgument : "value" "0"
var enemy_count = 0;

/// @DnDAction : YoYo Games.Common.Variable
/// @DnDVersion : 1
/// @DnDHash : 280881F3
/// @DnDInput : 2
/// @DnDArgument : "expr_1" "false"
/// @DnDArgument : "var" "enemies_remaining"
/// @DnDArgument : "var_1" "has_won"
enemies_remaining = 0;
has_won = false;

/// @DnDAction : YoYo Games.Loops.Repeat
/// @DnDVersion : 1
/// @DnDHash : 05DBF691
/// @DnDArgument : "times" "5"
repeat(5){	/// @DnDAction : YoYo Games.Common.Variable
	/// @DnDVersion : 1
	/// @DnDHash : 3B7597C9
	/// @DnDInput : 2
	/// @DnDParent : 05DBF691
	/// @DnDArgument : "expr" "1"
	/// @DnDArgument : "expr_relative" "1"
	/// @DnDArgument : "expr_1" "1"
	/// @DnDArgument : "expr_relative_1" "1"
	/// @DnDArgument : "var" "enemy_count"
	/// @DnDArgument : "var_1" "enemies_remaining"
	enemy_count += 1;
	enemies_remaining += 1;

	/// @DnDAction : YoYo Games.Instances.Create_Instance
	/// @DnDVersion : 1
	/// @DnDHash : 56B54BF2
	/// @DnDParent : 05DBF691
	/// @DnDArgument : "xpos" "50+enemy_count*200"
	/// @DnDArgument : "ypos" "80"
	/// @DnDArgument : "var" "new_enemy"
	/// @DnDArgument : "var_temp" "1"
	/// @DnDArgument : "objectid" "obj_enemy"
	/// @DnDSaveInfo : "objectid" "obj_enemy"
	var new_enemy = instance_create_layer(50+enemy_count*200, 80, "Instances", obj_enemy);

	/// @DnDAction : YoYo Games.Common.Apply_To
	/// @DnDVersion : 1
	/// @DnDHash : 792A35A4
	/// @DnDComment : Don't forget to set the Action Scope Expression$(13_10)to the variable name set in the Target field of$(13_10)the Create Instance Action Block above.  Use the$(13_10)drop down arrow in the upper right of the Apply$(13_10)To Action Block to set this...
	/// @DnDApplyTo : new_enemy
	/// @DnDParent : 05DBF691
	with(new_enemy) {
		/// @DnDAction : YoYo Games.Common.If_Expression
		/// @DnDVersion : 1
		/// @DnDHash : 6BE48318
		/// @DnDParent : 792A35A4
		/// @DnDArgument : "expr" "enemy_count % 2 == 0"
		if(enemy_count % 2 == 0){	/// @DnDAction : YoYo Games.Instances.Color_Sprite
			/// @DnDVersion : 1
			/// @DnDHash : 678C1A0F
			/// @DnDParent : 6BE48318
			/// @DnDArgument : "colour" "make_color_rgb(200, 25, 55) "
			/// @DnDArgument : "alpha" "false"
			image_blend = make_color_rgb(200, 25, 55)  & $ffffff;}
	
		/// @DnDAction : YoYo Games.Common.Else
		/// @DnDVersion : 1
		/// @DnDHash : 6C3FAE5C
		/// @DnDParent : 792A35A4
		else{	/// @DnDAction : YoYo Games.Instances.Color_Sprite
			/// @DnDVersion : 1
			/// @DnDHash : 32EE86CF
			/// @DnDParent : 6C3FAE5C
			/// @DnDArgument : "colour" "c_white"
			/// @DnDArgument : "alpha" "false"
			image_blend = c_white & $ffffff;}
	}}