using UnityEngine;
using System.Collections;


public class Enemy : MonoBehaviour
{
    public float          speed      = 1.5f;
    public int            pointValue = 10;
    public GameObject     explosionPrefab; 
    public SpriteRenderer spriteRenderer; 
    public Sprite         flashSprite;
    public Sprite         normalSprite;

    void Start()
    {
        normalSprite = spriteRenderer.sprite;

        // A little per-enemy variety on top of EnemySpawner's wave-based base 
        // speed, so a wave doesn't move as one perfectly uniform wall. Only for 
        // lane descent - formation members all share FormationController's one
        // moveSpeed instead, so this field goes unused while marching anyway.
        if (!FormationController.formationModeActive)
            { speed *= Random.Range(0.75f, 1.3f); }
    }

    void Update()
    {
        if (!GameManager.gameActive) return;
        if (FormationController.formationModeActive) return; // formation moves its own members

        transform.Translate(Vector3.down * speed * Time.deltaTime);
        if (transform.position.y < -6f)
        {
            GameManager.RegisterEscape();
            Destroy(gameObject);
        }
    }

    // Level 3 stretch: a short flicker (flash/normal/flash) across the same 0.1 seconds Level 1 already waits, instead of one flat swap.
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