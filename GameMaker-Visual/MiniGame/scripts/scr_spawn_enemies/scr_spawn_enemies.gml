/// @DnDAction : YoYo Games.Common.Function
/// @DnDVersion : 1
/// @DnDHash : 1F2B6254
/// @DnDComment : 
/// @DnDArgument : "funcName" "spawn_enemies"
function spawn_enemies() {	/// @DnDAction : YoYo Games.Common.Temp_Variable
	/// @DnDVersion : 1
	/// @DnDHash : 5D7A0B4E
	/// @DnDParent : 1F2B6254
	/// @DnDArgument : "var" "enemy_count"
	/// @DnDArgument : "value" "0"
	var enemy_count = 0;

	/// @DnDAction : YoYo Games.Loops.Repeat
	/// @DnDVersion : 1
	/// @DnDHash : 45FF1B09
	/// @DnDApplyTo : {obj_gamemanager}
	/// @DnDParent : 1F2B6254
	/// @DnDArgument : "times" "4 + wave"
	with(obj_gamemanager) {
	repeat(4 + wave){	/// @DnDAction : YoYo Games.Common.Variable
		/// @DnDVersion : 1
		/// @DnDHash : 3B80FFF3
		/// @DnDParent : 45FF1B09
		/// @DnDArgument : "expr" "1"
		/// @DnDArgument : "expr_relative" "1"
		/// @DnDArgument : "var" "enemy_count"
		enemy_count += 1;
	
		/// @DnDAction : YoYo Games.Instances.Create_Instance
		/// @DnDVersion : 1
		/// @DnDHash : 6677FEAC
		/// @DnDParent : 45FF1B09
		/// @DnDArgument : "xpos" "50 - (50 * obj_gamemanager.wave) + enemy_count * 200"
		/// @DnDArgument : "ypos" "80"
		/// @DnDArgument : "var" "new_enemy"
		/// @DnDArgument : "var_temp" "1"
		/// @DnDArgument : "objectid" "obj_enemy"
		/// @DnDSaveInfo : "objectid" "obj_enemy"
		var new_enemy = instance_create_layer(50 - (50 * obj_gamemanager.wave) + enemy_count * 200, 80, "Instances", obj_enemy);
	
		/// @DnDAction : YoYo Games.Common.Apply_To
		/// @DnDVersion : 1
		/// @DnDHash : 41D5F2CC
		/// @DnDComment : Don't forget to set the Action Scope Expression$(13_10)to the variable name set in the Target field of$(13_10)the Create Instance Action Block above.  Use the$(13_10)drop down arrow in the upper right of the Apply$(13_10)To Action Block to set this...
		/// @DnDApplyTo : new_enemy
		/// @DnDParent : 45FF1B09
		with(new_enemy) {
			/// @DnDAction : YoYo Games.Common.If_Expression
			/// @DnDVersion : 1
			/// @DnDHash : 44E8D874
			/// @DnDParent : 41D5F2CC
			/// @DnDArgument : "expr" "enemy_count % 2 == 0"
			if(enemy_count % 2 == 0){	/// @DnDAction : YoYo Games.Instances.Color_Sprite
				/// @DnDVersion : 1
				/// @DnDHash : 427B8A59
				/// @DnDParent : 44E8D874
				/// @DnDArgument : "colour" "make_color_rgb(200, 25, 55) "
				/// @DnDArgument : "alpha" "false"
				image_blend = make_color_rgb(200, 25, 55)  & $ffffff;}
		
			/// @DnDAction : YoYo Games.Common.Else
			/// @DnDVersion : 1
			/// @DnDHash : 2D3A20E6
			/// @DnDParent : 41D5F2CC
			else{	/// @DnDAction : YoYo Games.Instances.Color_Sprite
				/// @DnDVersion : 1
				/// @DnDHash : 61C6C1F5
				/// @DnDParent : 2D3A20E6
				/// @DnDArgument : "colour" "c_white"
				/// @DnDArgument : "alpha" "false"
				image_blend = c_white & $ffffff;}
		}}
	}}