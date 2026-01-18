import { createContext } from 'react';
import { Genre } from '@/lib/enums/genreEnums';

const GenreContext = createContext<Genre[]>([]);

export default GenreContext;
