extends Sprite2D

#region Global Variables
#
const TEX_DEFAULT    = preload("res://Assets/Sprites/player_circle_blue.png")
const TEX_RIGHT      = preload("res://Assets/Sprites/player_circle_yellow.png")
const TEX_LEFT       = preload("res://Assets/Sprites/player_circle_green.png")
const TEX_UP         = preload("res://Assets/Sprites/player_circle_red.png")
const TEX_DOWN       = preload("res://Assets/Sprites/player_circle_orange.png")
var   blink_time    := 0.0
var   spriteSize    := 60
var   bullet_offset := 40
var   fire_cooldown := 0.0
var   bullet_scene  := preload("res://Assets/Scenes/bullet.tscn")
#
#endregion

func _ready() -> void:
	pass

func _process(delta):
	#region Frame Variables
	#
	var moving   := Input.get_vector("ui_left", "ui_right", "ui_up", "ui_down")
	var viewPort := get_viewport_rect().size
	#
	#endregion
	
	#region Movement
	#
	position.x    = clamp(position.x, 0 + spriteSize, viewPort.x - spriteSize)
	position.y    = clamp(position.y, 0 + spriteSize, viewPort.y - spriteSize)
	
	if Input.is_action_pressed("ui_right"):
		position.x +=  200 * delta
		texture    = TEX_RIGHT
	if Input.is_action_pressed("ui_left"):
		position.x += -200 * delta
		texture     = TEX_LEFT
	if Input.is_action_pressed("ui_up"):
		position.y += -200 * delta
		texture     = TEX_UP
	if Input.is_action_pressed("ui_down"):
		position.y +=  200 * delta
		texture     = TEX_DOWN
	#
	#endregion
	
	#region Movement Animation and Sound
	#
	if !moving:
		texture    = TEX_DEFAULT
		visible    = true
		blink_time = 0.0
	else:
		blink_time += delta
		visible     = int(blink_time*6)%2 == 0
		if !$AudioStream_Walk.playing: 
			$AudioStream_Walk.play() 

	#
	#endregion
		
	#region Fire Control
	#
	fire_cooldown -= delta
	if     Input.is_action_just_pressed("ui_select") \
	   and fire_cooldown <= 0.0: 
		var b := bullet_scene.instantiate() 
		b.global_position = global_position
		b.global_position.y -= spriteSize + bullet_offset
		get_parent().add_child(b)
		fire_cooldown = 0.2

	#endregion
		
