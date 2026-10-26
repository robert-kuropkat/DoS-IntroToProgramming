draw_set_color(c_yellow);
draw_text(20, 30, "Lives: " + string(__dnd_lives));
draw_text(20, 50, "Score: " + string(__dnd_score));
draw_text(20, 70, "Enemy: " + string(obj_gamemanager.enemies_remaining));
draw_text(20, 90, "Wave:  " + string(wave) + " of " + string(total_waves));
draw_set_color(c_white);

if (has_won) 
{
	draw_set_color(c_yellow);
	draw_text(400, 50, "You Win! Press R to Restart");
	draw_set_color(c_white);
}
if (has_lost) 
{
	draw_set_color(c_yellow);
	draw_text(400, 50, "Game Over - Press R to Restart");
	draw_set_color(c_white);
}
if (wave_message_timer > 0) 
{ 
	draw_set_color(c_yellow);
	draw_text(400, 50, "Wave " + string(wave) + "!");
	draw_set_color(c_white);
}