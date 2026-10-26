using UnityEngine;

public class Enemy : MonoBehaviour
{
    public float enemySpeed = 1.5f;

    void Start()
    {
        enemySpeed = Random.Range(0.75f, 1.3f);
    }

    void Update()
    {
        if (!GameManager.gameActive) return;

        transform.Translate(Vector3.down * enemySpeed * Time.deltaTime);
        if (transform.position.y < -6f)
        {
            GameManager.RegisterEscape();
            Destroy(gameObject);
        }
    }
}