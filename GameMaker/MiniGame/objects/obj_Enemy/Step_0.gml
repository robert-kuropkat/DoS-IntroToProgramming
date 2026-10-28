if (!obj_gamemanager.game_active)  { exit; }

if      (  alarm[0] > 0 
        && alarm[0] mod 2 == 0
        ) 
                         { sprite_index = spr_enemy_flash; } 
else if ( alarm[0] > 0 ) { sprite_index = Sprite_Enemy;}

y += enemy_speed + ((obj_gamemanager.wave - 1) * 0.05);
if (y > room_height + 50) 
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