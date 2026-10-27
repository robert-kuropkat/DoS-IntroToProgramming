extends Sprite2D

func _ready():
	$Timer.timeout.connect(_on_timer_timeout)
	var tween = create_tween() 
	tween.tween_property(self, "scale", Vector2(1.5, 1.5), 0.2)

func _on_timer_timeout():
	queue_free()
