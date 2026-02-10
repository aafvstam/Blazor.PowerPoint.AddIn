/**
 * Type definitions for Blazor's DotNet global object
 */
interface DotNet {
  invokeMethodAsync<T>(assemblyName: string, methodName: string, ...args: any[]): Promise<T>;
  invokeMethod<T>(assemblyName: string, methodName: string, ...args: any[]): T;
}

declare const DotNet: DotNet;