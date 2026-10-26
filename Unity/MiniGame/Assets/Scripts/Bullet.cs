using UnityEngine;

public class Bullet : MonoBehaviour
{
    public float speed = 6f;
    public AudioClip hitSound;
    private float flashTimer;

    void Start()
    {
        GetComponent<SpriteRenderer>().color = Color.yellow;
        transform.localScale *= 1.3f;
        flashTimer = 0.1f;
    }

    void Update()
    {
        transform.Translate(Vector3.up * speed * Time.deltaTime);
        if (transform.position.y > 6f) Destroy(gameObject);

        if (flashTimer > 0f)
        {
            flashTimer -= Time.deltaTime;
            if (flashTimer <= 0f)
            {
                GetComponent<SpriteRenderer>().color = Color.white;
                transform.localScale /= 1.3f;
            }
        }
    }

    void OnTriggerEnter2D(Collider2D other)
    {
        if (other.CompareTag("Enemy"))
        {
            AudioSource.PlayClipAtPoint(hitSound, Camera.main.transform.position, 1.0f);
            GameManager.RegisterKill();
            Destroy(other.gameObject);
            Destroy(gameObject);
        }
    }
}