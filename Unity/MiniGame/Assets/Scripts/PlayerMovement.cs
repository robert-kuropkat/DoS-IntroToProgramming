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
                     private float          speed     = 3f;

    void Start() { sr.sprite = spriteDefault; }

    void Update()
    {
        float h = 0f;
        float v = 0f;
        if (Keyboard.current.rightArrowKey.isPressed) h =  1f;
        if (Keyboard.current.leftArrowKey.isPressed)  h = -1f;
        if (Keyboard.current.upArrowKey.isPressed)    v =  1f;
        if (Keyboard.current.downArrowKey.isPressed)  v = -1f;

        transform.Translate(new Vector3(h, v, 0) * speed * Time.deltaTime);
        transform.position = new Vector3( Mathf.Clamp(transform.position.x, -xBoundary, xBoundary)
                                        , Mathf.Clamp(transform.position.y, -yBoundary, yBoundary)
                                        , 0
                                        );

        bool moving = h != 0 || v != 0;
        if (moving)
        {
            blinkTime += Time.deltaTime;
            sr.enabled = ((int)(blinkTime * 6) % 2 == 0);
            if (!walkingBeep.isPlaying) walkingBeep.Play();

            if      (h > 0) sr.sprite = spriteRight;
            else if (h < 0) sr.sprite = spriteLeft;
            else if (v > 0) sr.sprite = spriteUp;
            else            sr.sprite = spriteDown;
        }
        else
        {
            sr.enabled = true;
            blinkTime  = 0f;
            walkingBeep.Stop();
            sr.sprite  = spriteDefault;
        }
    }
}