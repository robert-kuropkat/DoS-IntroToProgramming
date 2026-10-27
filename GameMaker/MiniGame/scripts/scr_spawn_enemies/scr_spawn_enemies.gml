function spawn_enemies(){
	
	for ( var enemy_count   = 0; 
              enemy_count   < 5; 
	          enemy_count++
	) 
    { 
		var new_enemy = instance_create_layer( 250 + enemy_count * 200
											 , 80
											 , "Instances"
											 , obj_enemy
											 );
		with (new_enemy) 
		{ 
			if (enemy_count % 2 == 0) { image_blend = make_color_rgb(200, 25, 55); }
			else                      { image_blend = c_white;                     }
		}
    }
}