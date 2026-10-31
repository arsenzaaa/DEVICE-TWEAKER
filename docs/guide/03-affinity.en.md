<!-- tag: CPU | summary: Interrupt Affinity policies, the AssignmentSetOverride mask, what IRQ Priority really does, Interrupt Steering, and the alternative "all cores except 0" method. -->

# Affinity, policies and interrupt priority

## Affinity

**Affinity** is the set of logical processors a driver's ISR / DPC (or a thread) may run on. For threads it is a hard rule: a priority-7 thread masked to core 1 preempts a priority-6 thread on core 1 even if core 2 is idle. An **IDEAL PROCESSOR** is only a preference: if that core is busy, the thread starts elsewhere.

For a device, affinity is set by a policy:

```text
HKLM\SYSTEM\CurrentControlSet\Enum\PCI\<device>\<instance>\Device Parameters\Interrupt Management\Affinity Policy
  DevicePolicy           REG_DWORD    policy (below)
  AssignmentSetOverride  REG_BINARY   CPU mask, SpecifiedProcessors only
  DevicePriority         REG_DWORD    IRQ priority (below)
```

```text
DevicePolicy  Windows policy                                In the program   What it does
0             IrqPolicyMachineDefault                       MachineDefault   Windows decides
1             IrqPolicyAllCloseProcessors                   AllClose         all CPUs close to the device (NUMA)
2             IrqPolicyOneCloseProcessor                    Single           one CPU close to the device
3             IrqPolicyAllProcessorsInMachine               All              any CPU
4             IrqPolicySpecifiedProcessors                  SpecCPU          only CPUs in the mask
5             IrqPolicySpreadMessagesAcrossAllProcessors    SpreadMessages   MSI-X messages on different CPUs
```

Only `SpecCPU` uses the checked mask; the other policies leave the choice to Windows. `SpreadMessages` only makes sense when the device and driver really use several MSI-X vectors; the policy does not create vectors.

## The mask

`AssignmentSetOverride` is a bit mask: bit N = logical processor N. CPUs 8 and 9:

```powershell
$mask = 0
@(8, 9) | ForEach-Object { $mask = $mask -bor (1 -shl $_) }
'{0:X16}' -f $mask    # 0000000000000300
```

In the registry the bytes go from least to most significant: `00 03 00 00 00 00 00 00`. DEVICE TWEAKER builds the mask from the checked processors and shows it as **Affinity Mask**. Processor group 0 is supported, i.e. up to 64 logical processors.

Clearing the CPU selection of an ordinary device writes `DevicePolicy = 0` and removes `AssignmentSetOverride`, returning the device to Windows.

## IRQ Priority

`DevicePriority` sets the IRQL of the interrupt within the device range:

```text
Undefined  = Normal
Low        = IRQL 3-4
Normal     = IRQL 5-11
High       = IRQL 12
```

It is **not** the game's priority, not a thread priority and not a DPC priority. Microsoft calls it a hint: Windows is not a real-time OS. IRQ priority affects which ISR may interrupt which when they overlap, and DPC generation; any ISR can still interrupt any other. The AUTO profile sets High on the devices it configures.

## Interrupt Steering

On multi-core systems Windows may redistribute interrupts itself (Interrupt Steering) and move them off parked cores. Steering events run in the DPC `ntoskrnl.exe!PpmCheckRun`, and redirection costs extra IPIs. With an explicit mask the interrupt goes straight to the right LAPIC. Steering itself can only be disabled through kernel variables, and the program does not do that.

## Devices that must not get a mask

Moving NVMe ISRs / DPCs is useless and dangerous: on some drives `DevicePolicy = SpecifiedProcessors` stops Windows from booting. So AUTO-OPTIMIZATION gives storage MSI and high priority but leaves affinity to Windows.

## If ISRs / DPCs do not move

If a trace shows the mask is not honoured (for example, a GPU on two cores but every ISR / DPC on one), set the hidden BIOS option **Interrupt Redirection Mode Selection = Round Robin** (via SCEWIN / GRUB):

```text
Fixed Priority  the first enabled APIC in the cluster
Round Robin     the next enabled APIC, round robin from the last selected one
Hash Vector     round robin starting from a hash of the vector number
```

## The alternative method

Another approach puts the GPU, XHCI and audio on **all cores except 0** (and except the network core). ISRs / DPCs then pick a core free of other interrupts and do not interrupt each other. The downsides:

- ISRs / DPCs still know nothing about threads and may interrupt any thread even when idle cores exist;
- ISRs / DPCs that run on a different core each time lose the cache advantage;
- only drivers that really gain from several cores (GPU, network, storage under load) should get them.

DEVICE TWEAKER uses the main method: fixed physical cores.

## Core 0

Core 0 receives more clock interrupts (it updates the system time), always runs the TIMER DPCs of NDIS, DXGMMS2, CLASSPNP and KBDHID, and SYMCRYPTK DPCs scale with the number of interrupts. AUTO-OPTIMIZATION keeps drivers off it when it can; with too few physical cores, core 0 is used last.

**Conclusion:** the `SpecCPU` policy with a mask on a specific physical core, High priority, NVMe left alone, result checked with a trace.
