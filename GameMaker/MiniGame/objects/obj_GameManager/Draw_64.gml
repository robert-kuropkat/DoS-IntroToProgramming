draw_set_color(c_yellow);
draw_text(20, 20, "Lives: " + string(__dnd_lives));
draw_text(20, 50, "Score: " + string(__dnd_score));
draw_set_color(c_white);

if (has_won) 
{
	draw_text(400, 50, "You Win!");
}
if (has_lost) 
{
	draw_text(400, 90, "Game Over");
}