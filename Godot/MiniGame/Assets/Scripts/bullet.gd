extends Sprite2D

var speed := 400.0
var flash_timer := 0.1

# Called when the node enters the scene tree for the first time.
func _ready() -> void:
	$AudioStream_Fire.play()
	modulate = Color.ORANGE 
	scale = Vector2(1.3, 1.3)

# Called every frame. 'delta' is the elapsed time since the previous frame.
func _process(delta: float) -> void:
	position.y -= speed * delta
	
	#region Remove Initial Fire Animation
	#
	if flash_timer > 0.0: 
		flash_timer -= delta 
		if flash_timer <= 0.0: 
			modulate = Color.WHITE 
			scale = Vector2(1, 1)
	#
	#endregion
	
	#region Cleanup
	#
	if position.y < -20:
		queue_free()
	#
	#endregion	


func _on_hitbox_area_entered(_area: Area2D) -> void:
		queue_free()
