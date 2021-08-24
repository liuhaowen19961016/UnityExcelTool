using UnityEngine;

public class Test : MonoBehaviour
{
    public Sheet1 data;

    private void Awake()
    {
        data.Init();
        string name = data.GetValue(1).name;
        Debug.Log(name);
    }
}
