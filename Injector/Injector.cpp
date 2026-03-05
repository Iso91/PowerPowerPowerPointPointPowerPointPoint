#include <Windows.h>
#include <TlHelp32.h>
#include <iostream>

int main() {
    HANDLE snap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    PROCESSENTRY32 pe = { sizeof(pe) };
    DWORD pid = 0;
    while (Process32Next(snap, &pe)) {
        if (wcscmp(pe.szExeFile, L"POWERPNT.EXE") == 0) {
            pid = pe.th32ProcessID;
            break;
        }
    }
    CloseHandle(snap);

    if (!pid) { std::cout << "PowerPoint not found\n"; return 1; }
    std::cout << "Found PowerPoint PID: " << pid << "\n";

    const char* dllPath = "C:\\Users\\timon\\source\\repos\\PPTHook\\x64\\Debug\\PPTHook.dll";

    HANDLE proc = OpenProcess(PROCESS_ALL_ACCESS, FALSE, pid);
    if (!proc) { std::cout << "OpenProcess failed: " << GetLastError() << "\n"; return 1; }

    LPVOID mem = VirtualAllocEx(proc, NULL, strlen(dllPath) + 1, MEM_COMMIT, PAGE_READWRITE);
    if (!mem) { std::cout << "VirtualAllocEx failed: " << GetLastError() << "\n"; return 1; }

    WriteProcessMemory(proc, mem, dllPath, strlen(dllPath) + 1, NULL);

    HANDLE thread = CreateRemoteThread(proc, NULL, 0,
        (LPTHREAD_START_ROUTINE)GetProcAddress(GetModuleHandle(L"kernel32.dll"), "LoadLibraryA"),
        mem, 0, NULL);
    if (!thread) { std::cout << "CreateRemoteThread failed: " << GetLastError() << "\n"; return 1; }

    WaitForSingleObject(thread, 5000);

    DWORD exitCode = 0;
    GetExitCodeThread(thread, &exitCode);
    std::cout << "Thread exit code (DLL base address): " << exitCode << "\n";
    if (exitCode == 0) std::cout << "DLL failed to load - probably blocked or wrong path\n";
    else std::cout << "DLL loaded successfully!\n";

    CloseHandle(thread);
    CloseHandle(proc);
    return 0;
}