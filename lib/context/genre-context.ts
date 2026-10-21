import { createContext } from 'react';
import { Genre } from '@/lib/enums/genre-enums';

const GenreContext = createContext<Genre[]>([]);

export default GenreContext;
