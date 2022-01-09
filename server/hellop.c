#include <stdio.h>  
#include "hello.h"

void HelloProc(unsigned char * pszString)

{

    printf("%s\n", pszString);

}

void Shutdown(void)

{
    RPC_STATUS status;
    status = RpcMgmtStopServerListening(NULL);

    // results into "server.obj : error LNK2019: ссылка на неразрешенный внешний символ _RPCServerUnregisterIf в функции _MIDL_user_free@4"
    // status = RPCServerUnregisterIf(NULL, NULL, FALSE);
}