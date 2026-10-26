function start_game(){
	with (obj_enemy) { instance_destroy(); } 
	
	obj_gamemanager.__dnd_score        = 0; 
	obj_gamemanager.__dnd_lives        = 3; 
	obj_gamemanager.enemies_remaining  = 5; 
	obj_gamemanager.wave               = 1;
	obj_gamemanager.wave_message_timer = 0;
	obj_gamemanager.has_won            = false; 
	obj_gamemanager.has_lost           = false; 
	obj_gamemanager.game_active        = true; 
	
	spawn_enemies();
	
}

function wave_cleared() 
{ 
	if (obj_gamemanager.wave >= obj_gamemanager.total_waves) 
	{ 
		obj_gamemanager.has_won             = true; 
		obj_gamemanager.game_active         = false; 
		audio_play_sound(sfx_win, 1, false);
	} else { 
		obj_gamemanager.wave               += 1; 
		obj_gamemanager.enemies_remaining   = 5; 
		obj_gamemanager.wave_message_timer  = game_get_speed(gamespeed_fps); 
		audio_play_sound(sfx_wave, 1, false);
		spawn_enemies(); 
	}
}
