//using Unity.VisualScripting.Antlr3.Runtime.Collections;
using UnityEngine;
using UnityEngine.InputSystem;

public class GameManager : MonoBehaviour
{
    public static int  enemiesRemaining     = 5;
    public static bool hasWon               = false;
    public static bool hasLost              = false;
    public static bool gameActive           = true;
    public static int  lives                = 3;
    public static int  score                = 0;
    public static int  wave                 = 1; 
    public static int  totalWaves           = 3; 
    public static bool nextWaveRequested    = false;
    public static float waveMessageTimeLeft = 0f;
    public static GameManager  instance;
    public        EnemySpawner spawner;
    public        AudioClip    lifeLostSound; 
    public        AudioClip    gameOverSound;
    public        AudioClip    gameStartSound;
    public        AudioClip    winSound;

    void Awake() { instance = this; }

    private void Start()
    {
        enemiesRemaining  = 5;
        lives             = 3;
        score             = 0;
        wave              = 1;
        totalWaves        = 3;
        nextWaveRequested = false;
        hasWon            = false;
        hasLost           = false;
        AudioSource.PlayClipAtPoint(instance.gameStartSound, Camera.main.transform.position);
    }

    void Update() 
    { 
        if (!gameActive && Keyboard.current.rKey.wasPressedThisFrame) { ResetGame(); }
        if (nextWaveRequested) 
        { 
            nextWaveRequested = false; 
            spawner.SpawnEnemies();
            waveMessageTimeLeft = 1f;
        }
        if (waveMessageTimeLeft > 0f) { waveMessageTimeLeft -= Time.deltaTime; }
    }

    public static void RegisterKill() 
    { 
        enemiesRemaining--;
        score += 10;
        if (enemiesRemaining <= 0) { WaveCleared(); }

        //if (enemiesRemaining <= 0 && !hasWon) 
        //{
        //    if (wave >= totalWaves)
        //    {
        //        hasWon = true;
        //        gameActive = false;
        //        AudioSource.PlayClipAtPoint(instance.winSound, Camera.main.transform.position);
        //    } else
        //    {
        //        wave++; 
        //        enemiesRemaining  = 5; 
        //        nextWaveRequested = true;
        //    }
        //}
    }

    public static void RegisterEscape()
    {
        lives--;
        enemiesRemaining--;
        //if (lives <= 0 && !hasLost)
        if (lives <= 0)
        {
            hasLost = true;
            gameActive = false;
            AudioSource.PlayClipAtPoint(instance.gameOverSound, Camera.main.transform.position);
            //} else if (lives > 0) { 
        } else if (enemiesRemaining <= 0) { WaveCleared();
        } else    { AudioSource.PlayClipAtPoint(instance.lifeLostSound, Camera.main.transform.position); }
    }

    public void ResetGame()
    {
        spawner.ClearEnemies();

        score               = 0; 
        lives               = 3; 
        wave                = 1;
        enemiesRemaining    = 5; 
        hasWon              = false; 
        hasLost             = false;
        nextWaveRequested   = false;
        waveMessageTimeLeft = 0f;
        gameActive          = true; 
        
        spawner.SpawnEnemies();
    }

    void OnGUI() 
    {
        GUIStyle bigStyle         = new GUIStyle(GUI.skin.label);   // Create a custom style
        bigStyle.fontSize         = 100;                            // Increase font size
        bigStyle.normal.textColor = Color.yellow;                   // Optional: change color

        if (hasWon)  GUI.Label( new Rect(10, 10, 1400, 300)
                              , "You Win! Press 'R' to replay"
                              , bigStyle
                              );
        if (hasLost) GUI.Label( new Rect(10, 10, 1400, 700)
                              , "Game Over : Press 'R' to replay"
                              , bigStyle
                              );
        if (waveMessageTimeLeft > 0f) 
                     GUI.Label( new Rect(10, 10, 1400, 700)
                              , "Wave " + wave + "!"
                              , bigStyle
                              );
        //if (waveMessageTimeLeft > 0f) { GUI.Label(new Rect(350, 150, 300, 40), "Wave " + wave + "!"); }
    }

    static void WaveCleared()
    {
        if (wave >= totalWaves)
        {
            hasWon = true; 
            gameActive = false; 
            AudioSource.PlayClipAtPoint(instance.winSound, Camera.main.transform.position); 
        } else { 
            wave++; 
            enemiesRemaining = 5; 
            nextWaveRequested = true; 
        }
    }


        }
