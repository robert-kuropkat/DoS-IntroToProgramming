extends Node2D

var enemy_scene := preload("res://Assets/Scenes/enemy.tscn")
var enemies_remaining := 0

# Called when the node enters the scene tree for the first time.
func _ready(): 
	return
	
# Called every frame. 'delta' is the elapsed time since the previous frame.
func _process(_delta: float) -> void:
	pass
	
func spawn_enemies(): 
	for enemy_count in range(5): 
		var new_enemy := enemy_scene.instantiate() 
		enemies_remaining += 1
		new_enemy.position = Vector2(150 + enemy_count * 200, 50) 
		if enemy_count % 2 == 0: 
			new_enemy.modulate = Color(0.8, 0.1, 0.2) 
		else: 
			new_enemy.modulate = Color(0.95, 0.4, 0.5)
		new_enemy.destroyed.connect($GameManager.on_enemy_destroyed)
		new_enemy.escaped.connect($GameManager.on_enemy_escaped)
		add_child(new_enemy)

#func _on_enemy_destroyed(): 
	#enemies_remaining -= 1 
	#if enemies_remaining <= 0: 
		#$WinLabel.show()
