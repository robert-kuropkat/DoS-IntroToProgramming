draw_text(20, 20, "Lives: " + string(__dnd_lives));
if (has_won) 
{
	draw_text(400, 50, "You Win!");
}
if (has_lost) 
{
	draw_text(400, 90, "Game Over");
}