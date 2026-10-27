using UnityEngine;
using System.Collections;


public class Enemy : MonoBehaviour
{
    public float          enemySpeed = 1.5f;
    public GameObject     explosionPrefab; 
    public SpriteRenderer spriteRenderer; 
    public Sprite         flashSprite;
    public Sprite         normalSprite;

    void Start()
    {
        enemySpeed   = Random.Range(0.75f, 1.3f);
        normalSprite = spriteRenderer.sprite;
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

    public IEnumerator HandleHit() 
    { 
        Instantiate(explosionPrefab, transform.position, Quaternion.identity);
        spriteRenderer.sprite = flashSprite;
        yield return new WaitForSeconds(0.03f);
        spriteRenderer.sprite = normalSprite;
        yield return new WaitForSeconds(0.03f);
        spriteRenderer.sprite = flashSprite;
        yield return new WaitForSeconds(0.04f);
        Destroy(gameObject); 
    }

}