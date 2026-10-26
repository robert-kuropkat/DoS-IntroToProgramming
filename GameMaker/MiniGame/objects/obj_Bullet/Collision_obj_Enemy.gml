audio_play_sound(sfx_enemyHit, 1, false);
obj_gamemanager.enemies_remaining--;
obj_gamemanager.__dnd_score += 10;

if (obj_gamemanager.enemies_remaining <= 0) { wave_cleared(); }

instance_destroy(other);
instance_destroy(self);