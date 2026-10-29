using UnityEngine;
using UnityEngine.InputSystem;

public class GameManager : MonoBehaviour
{
    public static int   enemiesRemaining    = 5;
    public static bool  hasWon              = false;
    public static bool  hasLost             = false;
    public static bool  gameActive          = false;
    public static bool  gameStarted         = false;
    public static int   lives               = 3;
    public static int   score               = 0;
    public static int   wave                = 1;
    public static int   totalWaves          = 3;
    public static bool  nextWaveRequested   = false;
    public static float waveMessageTimeLeft = 0f;
    public static int   startFrame          = -1; // the frame Space started the game - see Player.cs
    public static float winLoseStartTime    = -1f;

    public static GameManager  instance;
    public        EnemySpawner spawner;
    public        AudioClip    lifeLostSound;
    public        AudioClip    gameOverSound;
    public        AudioClip    gameStartSound;
    public        AudioClip    winSound;
    public        AudioClip    waveClip;

    void Awake() { instance = this; }

    void Start()
    {
        wave              = 1;
        enemiesRemaining  = 4 + wave;
        lives             = 3;
        score             = 0;
        totalWaves        = 3;
        winLoseStartTime  = -1f;
        nextWaveRequested = false;
        hasWon            = false;
        hasLost           = false;
        gameActive        = false;
        gameStarted       = false;
    }

    void Update()
    {
        // Level 3 stretch: title screen. Nothing else runs until the player 
        // presses Space - spawning is deliberately NOT in EnemySpawner.Start() 
        // any more, so the very first wave is only ever spawned from here.
        if (!gameStarted) 
        { 
            if (Keyboard.current.spaceKey.wasPressedThisFrame) 
            { 
                gameStarted = true; 
                gameActive  = true;
                startFrame  = Time.frameCount; // so Player.cs can tell this Space press apart from a fire input
                AudioSource.PlayClipAtPoint( instance.gameStartSound
                                           , Camera.main.transform.position
                                           ); 
                spawner.SpawnEnemies(); 
            } 
            return; 
        }

        if (!gameActive && Keyboard.current.rKey.wasPressedThisFrame) { ResetGame(); }

        if (nextWaveRequested)
        {
            nextWaveRequested   = false;
            spawner.SpawnEnemies();
            waveMessageTimeLeft = 1f;
        }

        if (waveMessageTimeLeft > 0f) { waveMessageTimeLeft -= Time.deltaTime; }
    }

    // Overload kept for any call site that doesn't care about a bonus-enemy 
    // point value - defaults to the Lesson 4 baseline of 10 points per kill.
    public static void RegisterKill() { RegisterKill(10); }
    
    public static void RegisterKill(int points)
    {
        enemiesRemaining--;
        score += points;
        if (enemiesRemaining <= 0) { WaveCleared(); }
    }

    public static void RegisterEscape()
    {
        lives--;
        enemiesRemaining--;
        if (lives <= 0)
        {
            winLoseStartTime = Time.time;
            hasLost          = true;
            gameActive       = false;
            AudioSource.PlayClipAtPoint(instance.gameOverSound, Camera.main.transform.position);
        }
        else if (enemiesRemaining <= 0)
        {
            WaveCleared();
        }
        else
        {
            AudioSource.PlayClipAtPoint(instance.lifeLostSound, Camera.main.transform.position);
        }
    }

    static void WaveCleared()
    {
        if (wave >= totalWaves)
        {
            winLoseStartTime = Time.time;
            hasWon           = true;
            gameActive       = false;
            AudioSource.PlayClipAtPoint(instance.winSound, Camera.main.transform.position);
        } else {
            wave++;
            enemiesRemaining  = 4 + wave;
            nextWaveRequested = true;
            AudioSource.PlayClipAtPoint(instance.waveClip, Camera.main.transform.position);
        }
    }

    public void ResetGame()
    {
        spawner.ClearEnemies();

        // Level 3 stretch: shields are pre-placed, not spawned, so a restart has to 
        // put each one back the way Shield.cs found it rather than re-instantiating it.
        foreach (   Shield shield 
                in FindObjectsByType<Shield>(FindObjectsInactive.Include)) 
                { shield.ResetShield(); }

        score               = 0;
        lives               = 3;
        wave                = 1;
        enemiesRemaining    = 4 + wave;
        waveMessageTimeLeft = 0f;
        winLoseStartTime    = -1f;
        hasWon              = false;
        hasLost             = false;
        nextWaveRequested   = false;
        gameActive          = true;
        startFrame          = Time.frameCount; // R is a fresh start too - keep Player.cs's guard honest

        spawner.SpawnEnemies();
    }

    // Level 3 stretch: title/game-over screen. Kept as OnGUI (immediate-mode), 
    // matching the win/lose banner this course already established in Lesson 10 - 
    // not moved to a Canvas, so the one existing message mechanism keeps doing all of it.
    void OnGUI()
    {
        GUIStyle bigStyle         = new GUIStyle(GUI.skin.label);
        bigStyle.fontSize         = 48;
        bigStyle.alignment        = TextAnchor.MiddleCenter;
        bigStyle.normal.textColor = Color.yellow;

        if (!gameStarted)           { GUI.Label(new Rect(10, 10, 1400, 300), "Kestrel Watch\nPress SPACE to begin your watch", bigStyle); return; }
        if (hasWon || hasLost)
        {
            float fadeInAlpha         = Mathf.Clamp01((Time.time - winLoseStartTime) / 0.5f);
            Color c                   = bigStyle.normal.textColor;
            bigStyle.normal.textColor = new Color(c.r, c.g, c.b, fadeInAlpha);
        }
        if (hasWon)                   GUI.Label(new Rect(10, 10, 1400, 300), "You Win! Press R to Restart"                   , bigStyle);
        if (hasLost)                  GUI.Label(new Rect(10, 10, 1400, 300), "Game Over - Press R to Restart"                , bigStyle);
        if (waveMessageTimeLeft > 0f) GUI.Label(new Rect(10, 10, 1400, 300), "Wave " + wave + "!"                            , bigStyle);
    }
}