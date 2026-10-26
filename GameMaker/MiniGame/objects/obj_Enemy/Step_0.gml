if (!obj_gamemanager.game_active)  { exit; }

y += enemy_speed;
if (y > room_height) 
{ 
	obj_gamemanager.__dnd_lives       -= 1; 
	obj_gamemanager.enemies_remaining -= 1;
	if (obj_gamemanager.__dnd_lives <= 0) 
	{ 
		obj_gamemanager.has_lost    = true; 
		obj_gamemanager.game_active = false;
		audio_play_sound(sfx_gameover, 1, false);
	} else 
		if (obj_gamemanager.enemies_remaining <= 0) { 
			wave_cleared();
		} else {
			audio_play_sound(sfx_lifelost, 1, false);
		}
	instance_destroy();
}