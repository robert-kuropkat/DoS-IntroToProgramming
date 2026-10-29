using UnityEngine;

public class FormationController : MonoBehaviour
{
    // Even waves march in formation; odd waves use Enemy.cs's own independent 
    // lane descent. This is what gives wave 2 a visibly different feel from 
    // waves 1 and 3, on top of the count/speed scaling both modes already share.
                     public  static bool  formationModeActive => GameManager.wave % 2 == 0;

    // xBoundary must stay wider than EnemySpawner's own spawn spread (half-width 
    // 3 + wave, so 5 on wave 2) or every member starts outside it and the formation 
    // free-falls on frame one instead of marching - keep a margin above that, not 
    // a number that merely matches it.
    [SerializeField] private        float moveSpeed = 1.5f; 
    [SerializeField] private        float stepDown  = 0.5f; 
    [SerializeField] private        float xBoundary = 8f; 
                     private        int   direction = 1;

    void Update()
    {
        if (  !formationModeActive 
           || !GameManager.gameActive
           ) return; 
        
        Enemy[] members = FindObjectsByType<Enemy>();
        if (members.Length == 0) return; 
        
        float dx = direction * moveSpeed * Time.deltaTime; 
        bool hitEdge = false; 
        foreach (Enemy e in members) 
        { 
            float nextX = e.transform.position.x + dx; 
            if (  nextX > xBoundary 
               || nextX < -xBoundary
               ) hitEdge = true; 
        }
        
        foreach (Enemy e in members) 
        { 
            if (hitEdge) e.transform.Translate(Vector3.down * stepDown); 
            else         e.transform.Translate(Vector3.right * dx); 
            
            if (e.transform.position.y < -6f) 
            { 
                GameManager.RegisterEscape(); 
                Destroy(e.gameObject); 
            } 
        }
        
        if (hitEdge) direction *= -1;
    }
}
