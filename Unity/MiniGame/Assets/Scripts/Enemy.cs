using UnityEditor.UI;
using UnityEngine;

public class Enemy : MonoBehaviour
{
    [SerializeField] private AudioClip enemyHit;
    [SerializeField] private float     enemySpeed = 0.75f;

    // Start is called once before the first execution of Update after the MonoBehaviour is created
    void Start()
    {
        enemySpeed = Random.Range(0.75f, 1.3f);
    }

    // Update is called once per frame
    void Update()
    {
        transform.Translate(Vector3.down * enemySpeed * Time.deltaTime); 
        if (transform.position.y < -6f) 
        {
            GameManager.RegisterEscape();
            Destroy(gameObject); 
        }
    }

    void OnTriggerEnter2D(Collider2D other)
    {
        if (other.CompareTag("Player"))
        {
            AudioSource.PlayClipAtPoint(enemyHit, transform.position);
            GameManager.RegisterKill();
            Destroy(gameObject);
        }
    }

}
