/// @DnDAction : YoYo Games.Common.Variable
/// @DnDVersion : 1
/// @DnDHash : 79D4670D
/// @DnDArgument : "expr" "1"
/// @DnDArgument : "expr_relative" "1"
/// @DnDArgument : "var" "obj_GameManager.enemies_remaining"
obj_GameManager.enemies_remaining += 1;

/// @DnDAction : YoYo Games.Random.Get_Random_Number
/// @DnDVersion : 1
/// @DnDHash : 3532895A
/// @DnDArgument : "var" "enemy_speed"
/// @DnDArgument : "min" "0.7"
/// @DnDArgument : "max" "1.3"
enemy_speed = (random_range(0.7, 1.3));