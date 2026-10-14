extends Sprite2D

signal destroyed
signal escaped

var enemy_speed := randf_range(50.0, 80.0)

# Called when the node enters the scene tree for the first time.
func _ready() -> void:
	pass

# Called every frame. 'delta' is the elapsed time since the previous frame.
func _process(delta: float) -> void:
	if not get_node("../GameManager").game_active: 
		return
	position.y += enemy_speed * delta 
	if position.y > 700: 
		escaped.emit()
		queue_free()


func _on_hitbox_area_entered(_area: Area2D) -> void:
	hide()
	$Hitbox.set_deferred("monitoring", false)
	$Hit.play()
	await $Hit.finished
	#area.get_parent().queue_free() 
	destroyed.emit()
	queue_free()
	
