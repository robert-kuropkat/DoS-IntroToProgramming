using UnityEngine;

public class EnemySpawner : MonoBehaviour
{
    [SerializeField] private GameObject enemyPrefab;
    [SerializeField] private int        enemyCount = 0;

    void Start()
    {
        SpawnEnemies();
    }

    public void ClearEnemies()
    {
        foreach (GameObject oldEnemy in GameObject.FindGameObjectsWithTag("Enemy")) { Destroy(oldEnemy); }
    }

    public void SpawnEnemies()
    {
        enemyCount = 4 + GameManager.wave;
        for ( int enemyIndex = 0; 
              enemyIndex < enemyCount; 
              enemyIndex++
            )
        {
            Vector3 pos         = new Vector3((-3f - GameManager.wave) + enemyIndex * 2f, 4f, 0f);
            GameObject newEnemy = Instantiate(enemyPrefab, pos, Quaternion.identity);
            if (enemyIndex % 2 == 0) { newEnemy.GetComponent<SpriteRenderer>().color = new Color(0.8f, 0.1f, 0.2f); }
            else                     { newEnemy.GetComponent<SpriteRenderer>().color = new Color(0.95f, 0.4f, 0.5f); }
        }
        GameManager.enemiesRemaining = enemyCount;
    }
}