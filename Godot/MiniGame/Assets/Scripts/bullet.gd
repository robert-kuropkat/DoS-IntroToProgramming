extends Sprite2D

var speed       := 400.0
var flash_timer := 0.1

func _ready() -> void:
	$AudioStream_Fire.play()
	modulate = Color.ORANGE 
	scale = Vector2(1.3, 1.3)

func _process(delta: float) -> void:
	position.y -= speed * delta
	
	#region Cleanup
	#
	if position.y < -20:
		queue_free()
	#
	#endregion	

	#region Remove Initial Fire Animation
	#
	if flash_timer > 0.0: 
		flash_timer -= delta 
		if flash_timer <= 0.0: 
			modulate = Color.WHITE 
			scale = Vector2(1, 1)
	#
	#endregion
	
