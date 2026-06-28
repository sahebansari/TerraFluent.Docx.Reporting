#if NETSTANDARD2_0
namespace System.Runtime.CompilerServices;

// Polyfill required by the compiler to support `init` accessors and positional
// records when targeting netstandard2.0, where this type doesn't ship in the BCL.
internal static class IsExternalInit
{
}
#endif
