function start_game(){
	with (obj_Enemy) { instance_destroy(); } 
	
	obj_GameManager.__dnd_score        = 0; 
	obj_GameManager.__dnd_lives        = 3; 
	obj_GameManager.enemies_remaining  = 0; 
	obj_GameManager.wave               = 1;
	obj_GameManager.wave_message_timer = 0;
	obj_GameManager.has_won            = false; 
	obj_GameManager.has_lost           = false; 
	obj_GameManager.game_active        = true; 
	
	spawn_enemies();
	
	//for ( var enemy_count = 0; 
    //      enemy_count < 5; 
	//      enemy_count++
	//) 
    //{ 
	//	var new_enemy = instance_create_layer( 50 + enemy_count * 200
	//										 , 80
	//										 , "Instances"
	//										 , obj_Enemy
	//										 );
	//	with (new_enemy) 
	//	{ 
	//		if (enemy_count % 2 == 0) { image_blend = make_color_rgb(200, 25, 55); }
	//		else                      { image_blend = c_white;                     }
	//	}
    //}
}

function wave_cleared() 
{ 
	if (obj_GameManager.wave >= obj_GameManager.total_waves) 
	{ 
		obj_GameManager.has_won             = true; 
		obj_GameManager.game_active         = false; 
	} else { 
		obj_GameManager.wave               += 1; 
		obj_GameManager.enemies_remaining   = 0; 
		obj_GameManager.wave_message_timer  = game_get_speed(gamespeed_fps); 
		spawn_enemies(); 
	}
}
