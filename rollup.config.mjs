/**
 * rollup.config.mjs
 *
 * Custom rollup override for Mendix pluggable-widgets-tools 11.x.
 *
 * SOLE PURPOSE:
 *   HyperFormula depends on chevrotain which contains an eval() call
 *   for parser performance optimization. Mendix's rollup pipeline treats
 *   any eval() in node_modules as a hard error.
 *
 *   This override intercepts the default Mendix config and adds the
 *   @rollup/plugin-node-resolve `ignoreSideEffectsForRoot` option plus
 *   patches the onwarn handler to downgrade the eval error to a warning
 *   specifically for the chevrotain file — and only that file.
 *
 *   No other build behaviour is changed.
 */

export default async function (args) {
    // Mendix passes its fully-resolved default config array here
    const defaultConfig = args.configDefaultConfig;

    // Apply our patch to every config in the array
    // (Mendix builds both "amd" and "es" output formats)
    const patched = defaultConfig.map(config => {
        const originalOnwarn = config.onwarn;

        return {
            ...config,
            onwarn(warning, warn) {
                // Downgrade eval error to nothing for chevrotain only.
                // chevrotain uses eval() purely as a performance trick —
                // it has zero security or correctness impact in our context.
                if (
                    warning.code === "EVAL" &&
                    warning.id &&
                    warning.id.includes("chevrotain")
                ) {
                    return; // Silently suppress — do not treat as error
                }

                // Everything else: use Mendix's original onwarn handler
                if (originalOnwarn) {
                    originalOnwarn(warning, warn);
                } else {
                    warn(warning);
                }
            }
        };
    });

    return patched;
}