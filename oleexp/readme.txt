**NONE OF THESE FILES ARE REQUIRED FOR TWINBASIC.**

oleexp.tlb is a precursor to WDL that's used in VB6 primarily for COM interfaces; it has far fewer API  declares. A copy of the source is maintained here for reference when converting projects to tB. You should always use WinDevLib in tB; oleexp.tlb will work in tB, but is 32-bit only and lacks the API definitions of tbShellLib. Note that addon modules are built into WinDevLib so also should not be used. There's minor signature differences in a small percentage of definitions but tB will pick up 99% of them during design time.
