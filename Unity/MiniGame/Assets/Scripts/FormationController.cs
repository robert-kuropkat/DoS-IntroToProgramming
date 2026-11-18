using System.Collections.Generic;
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
    //[SerializeField] private        float moveSpeed = 1.5f; 
    //[SerializeField] private        float stepDown  = 0.5f; 
    //[SerializeField] private        float xBoundary = 8f; 
    //                 private        int   direction = 1;
    [SerializeField] private float moveSpeed            = 4.5f;    // CHANGED - was 1.5f, read as too slow
    [SerializeField] private float maxMoveSpeed         = 7f;  // NEW - the ceiling moveSpeed can never escalate past
    [SerializeField] private float speedRampPerEdge     = 1.2f; // NEW - multiplier applied each time it hits an edge
    [SerializeField] private float stepDown             = 0.5f;
    [SerializeField] private float xBoundary            = 8f;
    [SerializeField] private float bonusSpeedMultiplier = 1.8f; // THIRD-PASS NEW - how much faster the bonus enemy marches than the rest of its group
                     private int   direction            = 1;
                     private float startingMoveSpeed; // NEW


    void Start() { startingMoveSpeed = moveSpeed; } // NEW

    void Update()
    {
        if (  !formationModeActive 
           || !GameManager.gameActive
           ) return;

        //Enemy[] members = FindObjectsByType<Enemy>();
        //if (members.Length == 0) return; 

        // NEW - isBonus enemies march on their own (see Enemy.cs); only
        // regular formation members are moved as one uniform group here.
        //Enemy[] allEnemies  = FindObjectsByType<Enemy>();
        //List<Enemy> members = new List<Enemy>();
        //foreach (Enemy e in allEnemies) { if (!e.isBonus) members.Add(e); }
        //if (members.Count == 0) return;
        Enemy[] allEnemies = FindObjectsByType<Enemy>();
        if (allEnemies.Length == 0) return;

        //float dx = direction * moveSpeed * Time.deltaTime; 
        bool hitEdge = false;
        //foreach (Enemy e in members) 
        foreach (Enemy e in allEnemies)
        {
            if (e.isBonus) continue; // THIRD-PASS NEW - its speed boost must not also speed up when the WHOLE group turns
            //float nextX = e.transform.position.x + dx; 
            float nextX = e.transform.position.x + direction * moveSpeed * Time.deltaTime;
            if (  nextX >  xBoundary 
               || nextX < -xBoundary
               ) hitEdge = true; 
        }

        //foreach (Enemy e in members) 
        foreach (Enemy e in allEnemies)
        {
            // if (hitEdge) e.transform.Translate(Vector3.down * stepDown); 
            // else         e.transform.Translate(Vector3.right * dx); 
            //if (hitEdge) { e.transform.Translate(Vector3.down * stepDown); }
            //else if (e.isBonus)   // THIRD-PASS NEW - faster than the group, clamped to the same boundary
            //{
            //    float bonusDx = direction * moveSpeed * bonusSpeedMultiplier * Time.deltaTime;
            //    float clampedNextX = Mathf.Clamp(e.transform.position.x + bonusDx, -xBoundary, xBoundary);
            //    e.transform.Translate(new Vector3(clampedNextX - e.transform.position.x, 0f, 0f));
            //} else { e.transform.Translate(Vector3.right * dx); }

            if (e.isBonus)
            {
                // Own direction, own speed, own edge check and step-down -
                // completely independent of the group above.
                float bonusDx = e.bonusDirection * moveSpeed * bonusSpeedMultiplier * Time.deltaTime;
                float nextX   = e.transform.position.x + bonusDx;
                if (  nextX >  xBoundary
                   || nextX < -xBoundary
                   )
                {
                    e.bonusDirection *= -1;
                    e.transform.Translate(Vector3.down * stepDown);
                } else { e.transform.Translate(Vector3.right * bonusDx); }
            } else if (hitEdge) {
                e.transform.Translate(Vector3.down * stepDown);
            } else {
                e.transform.Translate(Vector3.right * direction * moveSpeed * Time.deltaTime);
            }

            if (e.transform.position.y < -6f) 
            { 
                GameManager.RegisterEscape(); 
                Destroy(e.gameObject); 
            } 
        }

        //if (hitEdge) direction *= -1;
        if (hitEdge)
        {
            direction *= -1;
            //moveSpeed *= speedRampPerEdge; // NEW - each advance is a little faster than the last
            moveSpeed = Mathf.Min(moveSpeed * speedRampPerEdge, maxMoveSpeed); // NEW - same ramp, but capped
        }
    }

    // NEW - called by GameManager.ResetGame()
    public void ResetFormation()
    {
        moveSpeed = startingMoveSpeed;
        direction = 1;
    }

}
