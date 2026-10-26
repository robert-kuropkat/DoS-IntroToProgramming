using UnityEngine;
using UnityEngine.InputSystem;

public class Player : MonoBehaviour
{
    [SerializeField] private GameObject bulletPrefab;
    [SerializeField] private float fireCooldown = 0f;
    [SerializeField] private float bulletYOffset = 0.7f;
    [SerializeField] private AudioSource audioSrc;

    void Update()
    {
        fireCooldown -= Time.deltaTime;
        if (Keyboard.current.spaceKey.wasPressedThisFrame
           && fireCooldown <= 0f)
        {
            Instantiate(bulletPrefab
                       , transform.position + new Vector3(0, bulletYOffset, 0)
                       , Quaternion.identity
                       );
            fireCooldown = 0.2f;
            audioSrc.Play();
        }
    }
}