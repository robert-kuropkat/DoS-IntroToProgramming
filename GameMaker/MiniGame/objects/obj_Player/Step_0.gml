var moving =  keyboard_check(vk_right) 
           || keyboard_check(vk_left) 
		   || keyboard_check(vk_up) 
		   || keyboard_check(vk_down);
if (moving) { 
	          blink_timer += 1; 
			  visible = (blink_timer mod 20 < 10);
			  if (!audio_is_playing(sfx_blip)) audio_play_sound(sfx_blip, 1, false);
			} 
       else { 
		      visible = true; 
			  blink_timer = 0; 
			  audio_stop_sound(sfx_blip);	
			  sprite_index = Sprite_Blue;
			}

if (keyboard_check(vk_right)) { x += 4; sprite_index = Sprite_Yellow }
if (keyboard_check(vk_left))  { x -= 4; sprite_index = Sprite_Green }
if (keyboard_check(vk_up))    { y -= 4; sprite_index = Sprite_Red }
if (keyboard_check(vk_down))  { y += 4; sprite_index = Sprite_Orange }

if (fire_cooldown > 0) fire_cooldown -= 1;
if (  keyboard_check_pressed(vk_space) 
   && fire_cooldown <= 0) { 
	                        instance_create_layer(x+65, y, "Instances", obj_Bullet); 
							fire_cooldown = 12;
							audio_play_sound(sfx_shoot, 1, false);
						  }

x = clamp(x, 0, room_width  - 125);
y = clamp(y, 0, room_height - 125);