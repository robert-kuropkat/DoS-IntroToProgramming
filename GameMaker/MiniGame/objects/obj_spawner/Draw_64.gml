if (obj_GameManager.has_won) 
{ 
	draw_text(400, 50, "You Win! Enemies remaining: " 
	                 + string(obj_GameManager.enemies_remaining)
		     ); 
}

