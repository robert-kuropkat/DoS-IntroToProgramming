using UnityEngine;

public class Bullet : MonoBehaviour
{
    [SerializeField] private float speed = 6.0f;
    [SerializeField] private AudioSource audioSrc;
    [SerializeField] private float flashTimer;

    void Start()
    {
        GetComponent<SpriteRenderer>().color = Color.orange;
        transform.localScale *= 1.3f;
        flashTimer = 0.1f;
        audioSrc.Play();
    }

    void Update()
    {
        if (flashTimer > 0f)
        {
            flashTimer -= Time.deltaTime;
            if (flashTimer <= 0f)
            {
                GetComponent<SpriteRenderer>().color = Color.white;
                transform.localScale /= 1.3f;
            }
        }
        transform.Translate(Vector3.up * speed * Time.deltaTime);
        if (transform.position.y > 6f) Destroy(gameObject);
    }

    void OnTriggerEnter2D(Collider2D other)
    {
        if (other.CompareTag("Enemy"))
        {
            Destroy(gameObject);
        }
    }
}
