extends Node2D

var enemies_remaining := 5
var lives             := 3
var score             := 0

# Called when the node enters the scene tree for the first time.
func _ready() -> void:
	$GameStartSound.play()
	update_hud()
	
func update_hud():
	get_node("../LivesText").text = "Lives: " + str(lives)
	get_node("../ScoreText").text = "Score: " + str(score)

# Called every frame. 'delta' is the elapsed time since the previous frame.
func _process(_delta: float) -> void:
	pass

func on_enemy_destroyed(): 
	enemies_remaining -= 1 
	score += 10 
	update_hud()
	if enemies_remaining <= 0: 
		get_node("../WinText").show()
		$GameWinSound.play()
		
func on_enemy_escaped():
	lives -= 1 
	update_hud()
	#get_node("../LivesText").text = "Lives: " + str(lives)
	if lives <= 0: 
		get_node("../LoseText").show()
		$GameOverSound.play() 
	else: 
		$LifeLostSound.play()
