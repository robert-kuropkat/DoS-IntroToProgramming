audio_play_sound(sfx_enemyHit, 1, false);
obj_GameManager.enemies_remaining--;
//obj_spawner.enemies_remaining--;
//if (obj_spawner.enemies_remaining <= 0) { obj_spawner.has_won = true; }
instance_destroy(self);
