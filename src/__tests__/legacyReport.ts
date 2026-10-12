import { createReport as report } from '../index';
import { UserOptions } from '../types';

// The existing binary fixtures deliberately retain their legacy delimiters.
export const createReport = ((options: UserOptions, probe?: 'JS' | 'XML') =>
  report({ cmdDelimiter: '+++', ...options }, probe as 'JS')) as typeof report;
