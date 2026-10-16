if ( !game_active
   && keyboard_check_pressed(ord("R"))) { start_game(); }
if (wave_message_timer > 0)             { wave_message_timer -= 1; }