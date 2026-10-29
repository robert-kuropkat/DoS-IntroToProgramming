using UnityEngine;

public class PowerUp : MonoBehaviour
{
    public float fallSpeed = 1f;

    void Update()
    {
        transform.Translate(Vector3.down * fallSpeed * Time.deltaTime); 
        if (transform.position.y < -6f) Destroy(gameObject);
    }

    void OnTriggerEnter2D(Collider2D other) 
    { 
        if (!other.CompareTag("Player")) return; 
        
        Player player = other.GetComponent<Player>(); 
        if (player != null) player.ActivateRapidFire(5f); 
        Destroy(gameObject); 
    }
}
