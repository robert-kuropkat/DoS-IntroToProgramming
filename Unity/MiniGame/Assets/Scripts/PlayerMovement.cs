using System;
using UnityEngine;
using UnityEngine.InputSystem;

public class PlayerMovement : MonoBehaviour
{
    [SerializeField] private Sprite         spriteDefault, spriteLeft, spriteRight, spriteUp, spriteDown;
    [SerializeField] private SpriteRenderer sr;
    [SerializeField] private AudioSource    walkingBeep;
    [SerializeField] private float          xBoundary = 8f;
    [SerializeField] private float          yBoundary = 4.5f;
                     private float          blinkTime = 0f;
                     private float          h         = 0f;
                     private float          v         = 0f;
                     private float          speed     = 3f;
                     private bool           idle      = true;

    void Start()
    {
        sr.sprite = spriteDefault;
    }

    void Update()
    {
        h     = 0f;
        v     = 0f;

        if (Keyboard.current.rightArrowKey.isPressed) { h =  1f; sr.sprite = spriteRight; ImWalkingNow(); }
        if (Keyboard.current.leftArrowKey.isPressed)  { h = -1f; sr.sprite = spriteLeft ; ImWalkingNow(); }
        if (Keyboard.current.upArrowKey.isPressed)    { v =  1f; sr.sprite = spriteUp   ; ImWalkingNow(); }
        if (Keyboard.current.downArrowKey.isPressed)  { v = -1f; sr.sprite = spriteDown ; ImWalkingNow(); }

        idle = h == 0 && v == 0;
        if (idle) { sr.enabled = true; sr.sprite = spriteDefault; }

        transform.Translate(new Vector3(h, v, 0) * speed * Time.deltaTime);
        transform.position = new Vector3( Mathf.Clamp(transform.position.x, -xBoundary, xBoundary)
                                        , Mathf.Clamp(transform.position.y, -yBoundary, yBoundary)
                                        , 0);
    }

    void ImWalkingNow() 
    {
        blinkTime += Time.deltaTime;
        if (!walkingBeep.isPlaying) 
        { 
            walkingBeep.Play();
            sr.enabled = ((int)(blinkTime * 6) % 2 == 0);
        } 
    }

}
