extends Sprite2D

signal destroyed
signal escaped

const FLASH_TEXTURE  = preload("res://Assets/Sprites/enemy_invader_white.png")
const NORMAL_TEXTURE = preload("res://Assets/Sprites/enemy_invader_red.png")
var explosion_scene  = preload("res://Assets/Scenes/explosion.tscn")
var enemy_speed     := randf_range(50.0, 80.0)

func _ready() -> void:
	#$Hitbox.area_entered.connect(_on_hitbox_area_entered)
	pass

func _process(delta: float) -> void:
	if not get_node("../GameManager").game_active: 
		return
	position.y += enemy_speed * delta 
	if position.y > 700: 
		escaped.emit()
		queue_free()

func _on_hitbox_area_entered(area: Area2D) -> void:
	area.get_parent().queue_free() 
	$Hitbox.set_deferred("monitoring", false)
	$Hit.play()
	var explosion = explosion_scene.instantiate() 
	explosion.global_position = global_position 
	get_tree().current_scene.add_child(explosion) 
	texture = FLASH_TEXTURE
	await get_tree().create_timer(0.03).timeout 
	texture = NORMAL_TEXTURE
	await get_tree().create_timer(0.03).timeout 
	texture = FLASH_TEXTURE
	destroyed.emit()
	await get_tree().create_timer(0.04).timeout
	queue_free()
	
