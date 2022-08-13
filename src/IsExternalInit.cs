#if !NET6_0_OR_GREATER
using System.ComponentModel;

namespace System.Runtime.CompilerServices;

// https://stackoverflow.com/a/67687186
[EditorBrowsable(EditorBrowsableState.Never)]
public static class IsExternalInit
{
}
#endif
