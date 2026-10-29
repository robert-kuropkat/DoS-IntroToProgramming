using UnityEngine;
using UnityEngine.InputSystem;
using System.Collections;

public class Player : MonoBehaviour
{
    [SerializeField] private GameObject  bulletPrefab;
    [SerializeField] private float       fireCooldownMax = 0.2f;
    [SerializeField] private float       fireCooldown    = 0f;
    [SerializeField] private float       bulletYOffset   = 0.7f;
    [SerializeField] private AudioSource audioSrc;

    void Update()
    {
        fireCooldown -= Time.deltaTime;
        if (  Keyboard.current.spaceKey.wasPressedThisFrame
           && GameManager.gameActive 
           && Time.frameCount != GameManager.startFrame
           && fireCooldown    <= 0f
           )
        {
            Instantiate(bulletPrefab
                       , transform.position + new Vector3(0, bulletYOffset, 0)
                       , Quaternion.identity
                       );
            fireCooldown = fireCooldownMax;
            audioSrc.Play();
        }
    }

    public void ActivateRapidFire(float duration) { StartCoroutine(RapidFireRoutine(duration)); }

    private IEnumerator RapidFireRoutine(float duration) 
    { 
        float original  = fireCooldownMax; 
        fireCooldownMax = 0.05f; 
        yield return new WaitForSeconds(duration); 
        fireCooldownMax = original; 
    }
}