extends Node2D

var enemies_remaining := 5
var lives             := 3
var score             := 0
var wave              := 1
var total_waves       := 3
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
		#print ("Cleaning up enemies...")
		old_enemy.queue_free() 
	enemies_remaining = 5 
	lives             = 3 
	score             = 0 
	wave              = 1
	game_active = true 
	get_node("../WinText").hide() 
	get_node("../LoseText").hide() 
	get_node("../WaveAnnouncementText").hide()
	update_hud() 
	get_parent().spawn_enemies()

func update_hud():
	get_node("../LivesText").text = "Lives: " + str(lives)
	get_node("../ScoreText").text = "Score: " + str(score)
	get_node("../WaveText").text  = "Wave: "  + str(wave) + " of " + str(total_waves)

func on_enemy_destroyed(): 
	enemies_remaining -= 1 
	score += 10 
	update_hud()
	if enemies_remaining <= 0:
		wave_cleared()
	#if enemies_remaining <= 0: 
		#if wave >= total_waves:
			#game_active = false
			#get_node("../WinText").show()
			#$GameWinSound.play()
		#else: 
			#wave += 1 
			#enemies_remaining = 5 
			#update_hud() 
			#get_parent().call_deferred("spawn_enemies")
		
func on_enemy_escaped():
	lives -= 1 
	enemies_remaining -= 1
	update_hud()
	#get_node("../LivesText").text = "Lives: " + str(lives)
	if lives <= 0: 
		game_active = false
		get_node("../LoseText").show()
		$GameOverSound.play() 
	elif enemies_remaining <= 0: 
		wave_cleared()
	else: 
		$LifeLostSound.play()
		
func wave_cleared(): 
	if wave >= total_waves: 
		game_active = false 
		get_node("../WinText").text = "You Win! Press R to Restart" 
		get_node("../WinText").show()
		$GameWinSound.play() 
	else: 
		wave += 1 
		enemies_remaining = 5 
		update_hud() 
		get_parent().call_deferred("spawn_enemies") 
		announce_wave()
		
func announce_wave(): 
	#var banner = get_node("../WaveText") 
	var banner = get_node("../WaveAnnouncementText")
	banner.text = "Wave " + str(wave) + "!" 
	banner.show() 
	await get_tree().create_timer(1.0).timeout 
	banner.hide()
	
