using Unity.VisualScripting.Antlr3.Runtime.Collections;
using UnityEngine;

public class GameManager : MonoBehaviour
{
    public static int  enemiesRemaining = 5;
    public static bool hasWon           = false;
    public static bool hasLost          = false;
    public static int  lives            = 3;
    public static GameManager instance; 
    public        AudioClip   lifeLostSound; 
    public        AudioClip   gameOverSound;
    public        AudioClip   gameStartSound;
    public        AudioClip   winSound;

    void Awake() { instance = this; }

    private void Start()
    {
        enemiesRemaining = 5;
        lives            = 3;
        hasWon           = false;
        hasLost          = false;
        AudioSource.PlayClipAtPoint(instance.gameStartSound, Camera.main.transform.position);
    }
    public static void RegisterKill() 
    { 
        enemiesRemaining--;
        //if (enemiesRemaining <= 0) hasWon = true; 
        if (enemiesRemaining <= 0 && !hasWon) 
        { 
            hasWon = true; 
            AudioSource.PlayClipAtPoint(instance.winSound, Camera.main.transform.position); 
        }
    }

    public static void RegisterEscape()
    {
        lives--;
        //if (lives <= 0) hasLost = true;
        if (lives <= 0 && !hasLost) 
        { 
            hasLost = true; 
            AudioSource.PlayClipAtPoint(instance.gameOverSound, Camera.main.transform.position); 
        } else if (lives > 0) { 
            AudioSource.PlayClipAtPoint(instance.lifeLostSound, Camera.main.transform.position); 
        }
    }

    void OnGUI() 
    {
        GUIStyle bigStyle         = new GUIStyle(GUI.skin.label);   // Create a custom style
        bigStyle.fontSize         = 300;                            // Increase font size
        bigStyle.normal.textColor = Color.yellow;                   // Optional: change color

        GUI.Label(new Rect(10, 10, 1400, 700), "Lives: " + lives, bigStyle);
        if (hasWon)  GUI.Label( new Rect(10, 10, 1400, 300)
                              , "You Win!"
                              , bigStyle
                              );
//        if (hasLost) GUI.Label(new Rect(Screen.width / 2 - 50, 90, 1400, 700)
        if (hasLost) GUI.Label(new Rect(10, 10, 1400, 700)
                              , "Game Over"
                              , bigStyle
                              );
    }

}
