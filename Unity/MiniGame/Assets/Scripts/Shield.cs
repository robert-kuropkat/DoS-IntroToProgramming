using UnityEngine;

public class Shield : MonoBehaviour 
{ 
    public int            hitsRemaining = 3; 
    public SpriteRenderer spriteRenderer; 
    public Color[]        damageColors; // one entry per hit taken, brightest/least-damaged first
    private int           startingHits; 
    private Color         startingColor;

    void Start() 
    { 
        startingHits = hitsRemaining; 
        if (spriteRenderer != null) startingColor = spriteRenderer.color; 
    }

    void OnTriggerEnter2D(Collider2D other) 
    { 
        if (!other.CompareTag("Enemy")) return;
        if (!other.enabled)             return; // already being handled by a Bullet (or another Shield) this same frame

        other.enabled = false;
        Destroy(other.gameObject);
        GameManager.RegisterKill(0);           // no points - the shield earned this kill, not the player
        hitsRemaining--; 
        
        if (  spriteRenderer     != null 
           && damageColors       != null 
           && damageColors.Length > 0
           ) 
        { 
            int colorIndex       = Mathf.Clamp(damageColors.Length - hitsRemaining, 0, damageColors.Length - 1); 
            spriteRenderer.color = damageColors[colorIndex]; 
        } 
        
        if (hitsRemaining <= 0) Destroy(gameObject);
    }

    // Called by GameManager.ResetGame() - shields are pre-placed in the scene, 
    // not spawned from a prefab, so restoring one means putting it back the 
    // way Start() first found it rather than re-instantiating it.
    public void ResetShield()
    {
        hitsRemaining = startingHits; 
        if (spriteRenderer != null) spriteRenderer.color = startingColor; 
        gameObject.SetActive(true);
    }

}