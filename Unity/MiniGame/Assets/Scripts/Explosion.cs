using UnityEngine;

public class Explosion : MonoBehaviour
{
    void Start() { Destroy(gameObject, 0.3f); }

    private void Update() { transform.localScale += Vector3.one * Time.deltaTime * 2f; }
}
