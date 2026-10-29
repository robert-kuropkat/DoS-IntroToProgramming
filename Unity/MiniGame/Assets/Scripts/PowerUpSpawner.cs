using UnityEngine;

public class PowerUpSpawner : MonoBehaviour
{
    [SerializeField] private GameObject powerUpPrefab; 
    [SerializeField] private float      minInterval = 8f; 
    [SerializeField] private float      maxInterval = 15f; 
                     private float      timeUntilNext;

    void Start() { timeUntilNext = Random.Range(minInterval, maxInterval); }

    void Update()
    {
        if (!GameManager.gameActive) return; 
        
        timeUntilNext -= Time.deltaTime; 
        if (timeUntilNext <= 0f) 
        { 
            Vector3 pos = new Vector3(Random.Range(-3f, 3f), 4f, 0f); 
            Instantiate(powerUpPrefab, pos, Quaternion.identity); 
            timeUntilNext = Random.Range(minInterval, maxInterval); 
        }
    }
}
