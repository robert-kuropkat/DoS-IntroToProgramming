using UnityEngine;

public class EnemySpawner : MonoBehaviour
{
    [SerializeField] private GameObject enemyPrefab;

    public void ClearEnemies()
    {
        foreach (  GameObject oldEnemy 
                in GameObject.FindGameObjectsWithTag("Enemy")
                ) 
                { Destroy(oldEnemy); }
    }

    public void SpawnEnemies()
    {
        int count = 4 + GameManager.wave;
        for ( int enemyIndex = 0; 
              enemyIndex < count; 
              enemyIndex++
            )
        {
            Vector3 pos            = new Vector3((-3f - GameManager.wave) + enemyIndex * 2f, 4f, 0f);
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
            if (  enemyIndex           == count - 1 
               && GameManager.wave % 2 == 0
               ) 
            { 
                enemyScript.pointValue = 50; 
                enemyScript.speed     += 1f; 
                newEnemy.GetComponent<SpriteRenderer>().color = new Color(1f, 0.85f, 0.2f); 
            }
        }
        GameManager.enemiesRemaining = count;
    }
}