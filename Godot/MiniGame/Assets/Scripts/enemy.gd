extends Sprite2D

signal destroyed
signal escaped

var enemy_speed := randf_range(50.0, 80.0)

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
	hide()
	$Hitbox.set_deferred("monitoring", false)
	destroyed.emit()
	$Hit.play()
	await $Hit.finished
	queue_free()
	
