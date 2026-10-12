import { runUserJsAndGetRaw } from '../jsSandbox';
import { newContext } from '../processTemplate';
import { CreateReportOptions } from '../types';
import { CommandExecutionError, NullishCommandResultError } from '../errors';
import { newTextNode } from '../reportUtils';
import { customRunJs } from './customRunJs';

function context(overrides: Partial<CreateReportOptions> = {}) {
  return newContext({
    cmdDelimiter: ['{{', '}}'],
    literalXmlDelimiter: '||',
    processLineBreaks: true,
    additionalJsContext: {},
    failFast: true,
    rejectNullish: false,
    errorHandler: null,
    fixSmartQuotes: false,
    processLineBreaksAsNewText: false,
    indentXml: true,
    preserveSpace: true,
    compressionLevel: 1,
    ...overrides,
  });
}

it('executes directly by default, preserves assignments, and awaits snippet results', async () => {
  const ctx = context({
    additionalJsContext: { double: (x: number) => x * 2 },
  });
  await runUserJsAndGetRaw({ initial: 3 }, 'this.total = double(initial)', ctx);
  await expect(
    runUserJsAndGetRaw({}, 'Promise.resolve(total + 1)', ctx)
  ).resolves.toBe(7);
  await expect(runUserJsAndGetRaw({}, 'typeof process', ctx)).resolves.toBe(
    'object'
  );
  expect(ctx.jsSandbox?.__code__).toBeUndefined();
  expect(ctx.jsSandbox?.__result__).toBeUndefined();
});

it('awaits async runJs hooks and their results, and uses returned bindings in later commands', async () => {
  const runJs = jest.fn(async (input: Parameters<typeof customRunJs>[0]) => {
    const execution = await customRunJs(input);
    return {
      modifiedSandbox: { ...execution.modifiedSandbox, savedByHook: 10 },
      result: Promise.resolve(execution.result),
    };
  });
  const ctx = context({ runJs });
  await runUserJsAndGetRaw({ input: 2 }, 'total = input * 3', ctx);
  await expect(
    runUserJsAndGetRaw({}, 'total + savedByHook', ctx)
  ).resolves.toBe(16);
  expect(runJs).toHaveBeenCalledTimes(2);
  expect(runJs.mock.calls[1][0].sandbox).toMatchObject({
    __code__: 'total + savedByHook',
    total: 6,
    savedByHook: 10,
  });
  expect(runJs.mock.calls[1][0].ctx).toBe(ctx);
});

it('provides current loop values and additional context to runJs', async () => {
  const runJs = jest.fn(customRunJs);
  const ctx = context({ runJs, additionalJsContext: { suffix: '!' } });
  ctx.vars.item = 'Hello';
  ctx.loops.push({
    refNode: newTextNode(''),
    refNodeLevel: 1,
    varName: 'item',
    loopOver: ['Hello'],
    idx: 0,
  });
  await expect(
    runUserJsAndGetRaw({}, '$item + suffix + $idx', ctx)
  ).resolves.toBe('Hello!0');
  expect(runJs.mock.calls[0][0].sandbox).toMatchObject({
    $item: 'Hello',
    $idx: 0,
    suffix: '!',
  });
});

it('wraps a rejected custom result without falling back to direct execution', async () => {
  const failure = new Error('sandbox refused execution');
  const ctx = context({
    runJs: ({ sandbox }) => ({
      modifiedSandbox: sandbox,
      result: Promise.reject(failure),
    }),
  });
  await expect(
    runUserJsAndGetRaw({}, 'throw new Error("default executor ran")', ctx)
  ).rejects.toMatchObject({
    err: failure,
    command: 'throw new Error("default executor ran")',
  });
  await expect(runUserJsAndGetRaw({}, '1', ctx)).rejects.toBeInstanceOf(
    CommandExecutionError
  );
});

it('passes rejected async hooks to the error handler', async () => {
  const failure = new Error('sandbox failed');
  const errorHandler = jest.fn(async () => 'replacement');
  const ctx = context({
    runJs: async () => {
      throw failure;
    },
    errorHandler,
  });
  await expect(runUserJsAndGetRaw({}, 'command()', ctx)).resolves.toBe(
    'replacement'
  );
  expect(errorHandler).toHaveBeenCalledWith(failure, 'command()');
});

it('applies rejectNullish to a custom executor result', async () => {
  const ctx = context({
    rejectNullish: true,
    runJs: async ({ sandbox }) => ({
      modifiedSandbox: sandbox,
      result: Promise.resolve(null),
    }),
  });
  await expect(runUserJsAndGetRaw({}, 'value', ctx)).rejects.toBeInstanceOf(
    NullishCommandResultError
  );
});
