extends Node2D

var enemies_remaining := 5
var lives             := 3

# Called when the node enters the scene tree for the first time.
func _ready() -> void:
	$GameStartSound.play()
	get_node("../LivesText").text = "Lives: " + str(lives)


# Called every frame. 'delta' is the elapsed time since the previous frame.
func _process(_delta: float) -> void:
	pass

func on_enemy_destroyed(): 
	enemies_remaining -= 1 
	if enemies_remaining <= 0: 
		get_node("../WinText").show()
		$GameWinSound.play()
		
func on_enemy_escaped():
	lives -= 1 
	get_node("../LivesText").text = "Lives: " + str(lives)
	if lives <= 0: 
		get_node("../LoseText").show()
		$GameOverSound.play() 
	else: 
		$LifeLostSound.play()
