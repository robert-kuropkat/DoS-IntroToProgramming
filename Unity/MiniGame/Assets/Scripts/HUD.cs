using UnityEngine;
using UnityEngine.UI;

public class HUD : MonoBehaviour
{
    public Text scoreText; 
    public Text livesText;
    public Text waveText;

    void Update()
    {
        scoreText.text  = "Score: " + GameManager.score; 
        livesText.text  = "Lives: " + GameManager.lives;
        waveText.text   = "Wave: "  + GameManager.wave + " of " + GameManager.totalWaves;
    }
}
