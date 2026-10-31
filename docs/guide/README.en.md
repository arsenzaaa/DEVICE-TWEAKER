# Guide

[Project page](../../README.en.md) · [Русский](./README.md)

How it works in Windows and what the program does, device type by device type.

1. **[Interrupts: ISR, DPC and cores](./01-interrupts.en.md)**\
   How ISRs and DPCs work, why they outrank any game thread, and what happens when drivers share a core with GAME / RENDER / RHI.
2. **[MSI and MSI-X](./02-msi.en.md)**\
   Line versus message: why disabling MSI MODE adds extra and longer ISRs, why MessageNumberLimit should be removed, and how the program reads a device's real mode.
3. **[Affinity, policies and interrupt priority](./03-affinity.en.md)**\
   Interrupt Affinity policies, the AssignmentSetOverride mask, what IRQ Priority really does, Interrupt Steering, and the alternative "all cores except 0" method.
4. **[Graphics card](./04-gpu.en.md)**\
   Why a graphics card needs two different physical cores, how the NVIDIA driver differs from AMD, and what to do when ISRs / DPCs still pile up on one core.
5. **[USB controllers and IMOD](./05-usb-imod.en.md)**\
   How an xHCI controller works: interrupters, Event Rings and the IMOD register. Why a mouse gets a different interrupter on every connection, what IMOD to give audio, and how to spread devices across controllers.
6. **[Network: RSS, NDIS and NIC ITR](./06-network.en.md)**\
   NDIS versus NetAdapterCx, why a mask does not work on NDIS and RSS is needed instead, how many queues a game really uses, the NDIS thread and MMCSS, NIC ITR and the driver's own moderation.
7. **[CPU topology and ReservedCpuSets](./07-cpu-topology.en.md)**\
   SMT-SETs, P- and E-cores, CCDs and CPPC ranks: how to read the CPU list in the program. ReservedCpuSets, the mask format and two strategies for CPUs with two CCDs or E-cores.
8. **[Input: Polling Rate, Raw Input and Mouse Throttle](./08-input.en.md)**\
   The input path of a mouse, keyboard and controller on Windows 10 and 11, background-window throttling and RawMouseThrottleDuration, Raw Input Buffer and dedicated input threads, and what the Polling Rate measurement really measures.
9. **[Device power saving](./09-power.en.md)**\
   USB selective suspend, letting Windows turn a device off, and the Device Parameters values: what actually works, where it is stored, and how the program verifies the write.
10. **[Verification: traces and logs](./10-verification.en.md)**\
   How to check the result by facts rather than the FPS of one run: an xperf / WPR trace, where to look at ISRs and DPCs in WPA and MXA, what to send in a hardware report and how to read the program log.
11. **[DTIMOD.sys, KDU and startup](./11-driver.en.md)**\
   Why the program has its own driver, where it may write, why KDU loads it and what antivirus software, HVCI and anti-cheats think about that. How IMOD / NIC ITR startup works without the EXE.
