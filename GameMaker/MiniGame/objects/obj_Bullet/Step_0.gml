y -= 8;

if (flash_timer > 0) { 
	                   flash_timer -= 1; 
					   if (flash_timer <= 0) { 
						                       image_blend = c_white; 
											   image_xscale = 1; 
											   image_yscale = 1; 
											 }
				     }

if (y < 0) instance_destroy();
