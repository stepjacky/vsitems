// APClusterDllMedium.cpp : ���� DLL Ӧ�ó���ĵ���������
//

#include "stdafx.h"
 #include <stdio.h>
#include "jni.h"

 #ifdef __cplusplus
 extern "C" {
 #endif
 
 typedef int*  (__stdcall *APCLUSTER32)(double*, unsigned int, bool);
 
 JNIEXPORT jintArray JNICALL Java_APCluster_CallAPClusterDll
   (JNIEnv *env, jobject _obj, jint _arg_int, jdoubleArray _arg_doublearray, jboolean _arg_boolean)
 {
     
     return NULL;
 }
 
 #ifdef __cplusplus
 }
 #endif

