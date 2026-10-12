import vm from 'vm';
import { RunJSFunc } from '../types';

// A caller can choose a VM in its hook; the library has no VM dependency.
export const customRunJs: RunJSFunc = ({ sandbox }) => {
  const context = vm.createContext(sandbox);
  const result = new vm.Script(sandbox.__code__ ?? '').runInContext(context);
  return { modifiedSandbox: sandbox, result };
};
