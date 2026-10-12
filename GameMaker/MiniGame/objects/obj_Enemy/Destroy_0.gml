if (obj_GameManager.enemies_remaining <= 0) 
{ 
	obj_GameManager.has_won = true; 
	audio_play_sound(sfx_win, 1, false);
}