audio_play_sound(sfx_enemyHit, 1, false);
instance_create_layer(other.x, other.y, "Instances", obj_explosion);
other.sprite_index = spr_enemy_flash;
other.image_blend  = c_white;
other.alarm[0]     = 6;
obj_gamemanager.enemies_remaining--;
obj_gamemanager.__dnd_score += 10;

if (obj_gamemanager.enemies_remaining <= 0) { wave_cleared(); }

instance_destroy(self);