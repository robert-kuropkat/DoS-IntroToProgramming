using UnityEngine;
using System.Collections;
//using static UnityEditor.PlayerSettings;

public class EnemySpawner : MonoBehaviour
{
    [SerializeField] private GameObject enemyPrefab;
    [SerializeField] private float      bonusSpawnDelay = 0.75f;
                     private Coroutine  bonusCoroutine; // NEW - tracked so ClearEnemies() can cancel a still-pending bonus spawn

    public void ClearEnemies()
    {
        if (bonusCoroutine != null) // NEW - cancel a pending bonus spawn along with every actual GameObject
        {
            StopCoroutine(bonusCoroutine);
            bonusCoroutine = null;
        }
        foreach (  GameObject oldEnemy 
                in GameObject.FindGameObjectsWithTag("Enemy")
                ) 
                { Destroy(oldEnemy); }
    }

    public void SpawnEnemies()
    {
        int count          = 4 + GameManager.wave;
        bool hasBonusEnemy = GameManager.wave % 2 == 0;
        for ( int enemyIndex = 0; 
              enemyIndex < count; 
              enemyIndex++
            )
        {
            //Vector3 pos            = new Vector3((-3f - GameManager.wave) + enemyIndex * 2f, 4f, 0f);
            // THIRD-PASS NEW - the bonus enemy (last slot, even waves only)
            // doesn't spawn here anymore - see SpawnBonusEnemy() below.
            if (hasBonusEnemy && enemyIndex == count - 1) continue;

            Vector3    pos         = new Vector3((-3f - GameManager.wave) + enemyIndex * 2f, 3.5f, 0f); // THIRD-PASS: row Y nudged from 4f to 3.5f, a touch more room underneath
            GameObject newEnemy    = Instantiate(enemyPrefab, pos, Quaternion.identity);

            Enemy enemyScript      = newEnemy.GetComponent<Enemy>();
            enemyScript.speed      = 1.5f + (GameManager.wave - 1) * 0.5f;
            enemyScript.pointValue = Random.Range(5, 16);  // Lesson 9 Level 3: scoring variety

            // Lesson 5 Level 3: alternating striped-row coloring, by spawn slot - 
            // deliberately simple shades so it reads as a pattern, not noise.
            newEnemy.GetComponent<SpriteRenderer>().color = ( enemyIndex % 2 == 0)
                                                            ? new Color(0.8f, 0.1f, 0.2f)
                                                            : new Color(0.95f, 0.4f, 0.5f);

            // Level 3 stretch: a bonus enemy - the last slot in every even-numbered 
            // wave, worth more points and a little faster, using the exact same 
            // Enemy prefab rather than a second one. Overrides both the random 
            // point value and the striping above, same as it already overrides speed.
            //if (enemyIndex == count - 1
            //   && GameManager.wave % 2 == 0
            //   )
            //{
            //    enemyScript.pointValue = 50;
            //    enemyScript.speed           += 1f;                      // THIRD-PASS: removed - speed is no longer read for the bonus enemy; FormationController now owns its pace
            //    enemyScript.isBonus = true;                      // NEW - exempt from FormationController's uniform movement
            //    newEnemy.transform.position += new Vector3(0f, 1.5f, 0f); // NEW - spawn above the row
            //    newEnemy.GetComponent<SpriteRenderer>().color = new Color(1f, 0.85f, 0.2f);
            //}
        }
        if (hasBonusEnemy)
        {
            Vector3 bonusPos = new Vector3((-3f - GameManager.wave) + (count - 1) * 2f, 3.5f, 0f);
            //StartCoroutine(SpawnBonusEnemy(bonusPos)); // THIRD-PASS NEW
            bonusCoroutine = StartCoroutine(SpawnBonusEnemy(bonusPos)); // NEW - store the reference so ClearEnemies() can cancel it
        }
        GameManager.enemiesRemaining = count;
    }

    // THIRD-PASS NEW - spawning it a short delay after the rest of the row,
    // at the SAME row height (not above it), is what makes it genuinely
    // start behind the formation's action for a moment, without needing a
    // permanent spatial offset that never closes.
    private IEnumerator SpawnBonusEnemy(Vector3 pos)
    {
        yield return new WaitForSeconds(bonusSpawnDelay);
        GameObject newEnemy    = Instantiate(enemyPrefab, pos, Quaternion.identity);
        Enemy enemyScript      = newEnemy.GetComponent<Enemy>();
        enemyScript.pointValue = 50;
        enemyScript.isBonus    = true;
        // THIRD-PASS NEW - see the separate color note below for why
        SpriteRenderer sr      = newEnemy.GetComponent<SpriteRenderer>();
        sr.sprite              = enemyScript.flashSprite; // the white sprite already wired on every Enemy prefab
        sr.color               = new Color(1f, 0.85f, 0.2f); // gold
        bonusCoroutine         = null; // NEW - finished normally, nothing left to cancel
    }

}