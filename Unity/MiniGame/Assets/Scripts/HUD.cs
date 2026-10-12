using UnityEngine;
using UnityEngine.UI;

public class HUD : MonoBehaviour
{
    public Text scoreText; 
    public Text livesText;
 
    void Update()
    {
        scoreText.text = "Score: " + GameManager.score; 
        livesText.text = "Lives: " + GameManager.lives;
    }
}
