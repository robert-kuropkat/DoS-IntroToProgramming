if (!obj_GameManager.game_active)  { exit; }

y += enemy_speed;
if (y > room_height) 
{ 
	//show_debug_message("Enemy Escaped!");
	obj_GameManager.__dnd_lives -= 1; 
	if (obj_GameManager.__dnd_lives <= 0) 
	{ 
		obj_GameManager.has_lost    = true; 
		obj_GameManager.game_active = false;
		audio_play_sound(sfx_gameover, 1, false);
	} else {
		audio_play_sound(sfx_lifelost, 1, false);
	}
	instance_destroy();
}