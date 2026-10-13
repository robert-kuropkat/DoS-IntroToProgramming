/// @DnDAction : YoYo Games.Common.Function
/// @DnDVersion : 1
/// @DnDHash : 60377EA3
/// @DnDComment : 
/// @DnDArgument : "funcName" "start_game"
function start_game() {	/// @DnDAction : YoYo Games.Common.Apply_To
	/// @DnDVersion : 1
	/// @DnDHash : 07A24CA2
	/// @DnDApplyTo : {obj_enemy}
	/// @DnDParent : 60377EA3
	with(obj_enemy) {
		/// @DnDAction : YoYo Games.Instances.Destroy_Instance
		/// @DnDVersion : 1
		/// @DnDHash : 4F1A6E17
		/// @DnDApplyTo : {obj_enemy}
		/// @DnDParent : 07A24CA2
		with(obj_enemy) instance_destroy();
	}

	/// @DnDAction : YoYo Games.Instance Variables.Set_Score
	/// @DnDVersion : 1
	/// @DnDHash : 48973A6D
	/// @DnDApplyTo : {obj_GameManager}
	/// @DnDParent : 60377EA3
	with(obj_GameManager) {
	__dnd_score = real(0);
	}

	/// @DnDAction : YoYo Games.Instance Variables.Set_Lives
	/// @DnDVersion : 1
	/// @DnDHash : 78006E9C
	/// @DnDApplyTo : {obj_GameManager}
	/// @DnDParent : 60377EA3
	/// @DnDArgument : "lives" "3"
	with(obj_GameManager) {
	__dnd_lives = real(3);
	}

	/// @DnDAction : YoYo Games.Common.Variable
	/// @DnDVersion : 1
	/// @DnDHash : 446D1A1F
	/// @DnDInput : 4
	/// @DnDApplyTo : {obj_GameManager}
	/// @DnDParent : 60377EA3
	/// @DnDArgument : "expr_1" "false"
	/// @DnDArgument : "expr_2" "false"
	/// @DnDArgument : "expr_3" "true"
	/// @DnDArgument : "var" "enemies_remaining"
	/// @DnDArgument : "var_1" "has_won"
	/// @DnDArgument : "var_2" "has_lost"
	/// @DnDArgument : "var_3" "game_active"
	with(obj_GameManager) {
	enemies_remaining = 0;
	has_won = false;
	has_lost = false;
	game_active = true;
	
	}

	/// @DnDAction : YoYo Games.Common.Temp_Variable
	/// @DnDVersion : 1
	/// @DnDHash : 21E3DBBD
	/// @DnDParent : 60377EA3
	/// @DnDArgument : "var" "enemy_count"
	/// @DnDArgument : "value" "0"
	var enemy_count = 0;

	/// @DnDAction : YoYo Games.Loops.Repeat
	/// @DnDVersion : 1
	/// @DnDHash : 1E975330
	/// @DnDParent : 60377EA3
	/// @DnDArgument : "times" "5"
	repeat(5){	/// @DnDAction : YoYo Games.Common.Variable
		/// @DnDVersion : 1
		/// @DnDHash : 59BA209B
		/// @DnDParent : 1E975330
		/// @DnDArgument : "expr" "1"
		/// @DnDArgument : "expr_relative" "1"
		/// @DnDArgument : "var" "enemy_count"
		enemy_count += 1;
	
		/// @DnDAction : YoYo Games.Instances.Create_Instance
		/// @DnDVersion : 1
		/// @DnDHash : 3F3F7055
		/// @DnDParent : 1E975330
		/// @DnDArgument : "xpos" "50+enemy_count*200"
		/// @DnDArgument : "ypos" "80"
		/// @DnDArgument : "var" "new_enemy"
		/// @DnDArgument : "var_temp" "1"
		/// @DnDArgument : "objectid" "obj_enemy"
		/// @DnDSaveInfo : "objectid" "obj_enemy"
		var new_enemy = instance_create_layer(50+enemy_count*200, 80, "Instances", obj_enemy);
	
		/// @DnDAction : YoYo Games.Common.Apply_To
		/// @DnDVersion : 1
		/// @DnDHash : 7AA95AFF
		/// @DnDComment : Don't forget to set the Action Scope Expression$(13_10)to the variable name set in the Target field of$(13_10)the Create Instance Action Block above.  Use the$(13_10)drop down arrow in the upper right of the Apply$(13_10)To Action Block to set this...
		/// @DnDApplyTo : new_enemy
		/// @DnDParent : 1E975330
		with(new_enemy) {
			/// @DnDAction : YoYo Games.Common.If_Expression
			/// @DnDVersion : 1
			/// @DnDHash : 416BE97F
			/// @DnDParent : 7AA95AFF
			/// @DnDArgument : "expr" "enemy_count % 2 == 0"
			if(enemy_count % 2 == 0){	/// @DnDAction : YoYo Games.Instances.Color_Sprite
				/// @DnDVersion : 1
				/// @DnDHash : 40A8482B
				/// @DnDParent : 416BE97F
				/// @DnDArgument : "colour" "make_color_rgb(200, 25, 55) "
				/// @DnDArgument : "alpha" "false"
				image_blend = make_color_rgb(200, 25, 55)  & $ffffff;}
		
			/// @DnDAction : YoYo Games.Common.Else
			/// @DnDVersion : 1
			/// @DnDHash : 39315D1A
			/// @DnDParent : 7AA95AFF
			else{	/// @DnDAction : YoYo Games.Instances.Color_Sprite
				/// @DnDVersion : 1
				/// @DnDHash : 5FCB6FE7
				/// @DnDParent : 39315D1A
				/// @DnDArgument : "colour" "c_white"
				/// @DnDArgument : "alpha" "false"
				image_blend = c_white & $ffffff;}
		}}}