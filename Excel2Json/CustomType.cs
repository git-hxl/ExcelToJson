
using System;
#if UNITY_5_3_OR_NEWER
    using UnityEngine; 
#endif


namespace Excel2Json
{
#if !UNITY_5_3_OR_NEWER
    public class Vector2
    {
        public float x;
        public float y;
        public Vector2(float x, float y)
        {
            this.x = x;
            this.y = y;
        }
    }

    public class Vector3
    {
        public float x;
        public float y;
        public float z;
        public Vector3(float x, float y, float z)
        {
            this.x = x;
            this.y = y;
            this.z = z;
        }
    }

    public class Vector4
    {
        public float x;
        public float y;
        public float z;
        public float w;
        public Vector4(float x, float y, float z, float w)
        {
            this.x = x; this.y = y;
            this.z = z; this.w = w;
        }
    }

    public class Color
    {
        public float r; public float g; public float b; public float a;

        public Color(float r, float g, float b, float a)
        {
            this.r = r;
            this.g = g;
            this.b = b;
            this.a = a;
        }
    }
#endif

    public class CusTomType
    {
        /// <summary>
        /// 根据名称返回类型
        /// </summary>
        /// <param name="typeName"></param>
        /// <returns></returns>
        public static Type GetTypeByString(string typeName)
        {
            switch (typeName.ToLower())
            {
                case "uint":
                    return typeof(uint);
                case "int":
                    return typeof(int);
                case "long":
                    return typeof(long);
                case "float":
                    return typeof(float);
                case "double":
                    return typeof(double);
                case "bool":
                    return typeof(bool);
                case "string":
                    return typeof(string);

                case "vector2":
                    return typeof(Vector2);
                case "vector3":
                    return typeof(Vector3);
                case "vector4":
                    return typeof(Vector4);
                case "color":
                    return typeof(Color);

                case "int[]":
                    return typeof(int[]);
                case "long[]":
                    return typeof(long[]);
                case "float[]":
                    return typeof(float[]);
                case "double[]":
                    return typeof(double[]);
                case "bool[]":
                    return typeof(bool[]);
                case "string[]":
                    return typeof(string[]);

                case "vector2[]":
                    return typeof(Vector2[]);
                case "vector3[]":
                    return typeof(Vector3[]);
                case "vector4[]":
                    return typeof(Vector4[]);
                case "color[]":
                    return typeof(Color[]);

                case "date":
                case "datetime":
                    return typeof(DateTime);
                default:
                    return typeof(string);
            }
        }
    }

}
