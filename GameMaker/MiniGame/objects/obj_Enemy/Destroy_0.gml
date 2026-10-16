if (obj_GameManager.enemies_remaining <= 0) 
{ 
	if (obj_GameManager.wave >= obj_GameManager.total_waves) {
		obj_GameManager.has_won     = true; 
		obj_GameManager.game_active = false;
		audio_play_sound(sfx_win, 1, false);
	} else {
		//obj_GameManager.wave++;
		//obj_GameManager.enemies_remaining = 0;
		//spawn_enemies();
	}
	wave_cleared();
}