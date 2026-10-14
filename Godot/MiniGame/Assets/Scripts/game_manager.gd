extends Node2D

var enemies_remaining := 5
var lives             := 3
var score             := 0
var game_active       := true;

# Called when the node enters the scene tree for the first time.
func _ready() -> void:
	$GameStartSound.play()
	start_game.call_deferred()
	# update_hud()

# Called every frame. 'delta' is the elapsed time since the previous frame.
func _process(_delta: float) -> void:
	pass

func _input(event): 
	if     event is InputEventKey \
	   and event.pressed \
	   and event.keycode == KEY_R: 
		if not game_active: 
			start_game()
	
func start_game(): 
	for old_enemy in get_tree().get_nodes_in_group("enemies"): 
		print ("Cleaning up enemies...")
		old_enemy.queue_free() 
	enemies_remaining = 5 
	lives             = 3 
	score             = 0 
	game_active = true 
	get_node("../WinText").hide() 
	get_node("../LoseText").hide() 
	update_hud() 
	get_parent().spawn_enemies()

func update_hud():
	get_node("../LivesText").text = "Lives: " + str(lives)
	get_node("../ScoreText").text = "Score: " + str(score)

func on_enemy_destroyed(): 
	enemies_remaining -= 1 
	score += 10 
	update_hud()
	if enemies_remaining <= 0: 
		game_active = false
		get_node("../WinText").show()
		$GameWinSound.play()
		
func on_enemy_escaped():
	lives -= 1 
	update_hud()
	#get_node("../LivesText").text = "Lives: " + str(lives)
	if lives <= 0: 
		game_active = false
		get_node("../LoseText").show()
		$GameOverSound.play() 
	else: 
		$LifeLostSound.play()
