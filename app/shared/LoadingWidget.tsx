// library imports
import CircularProgress from '@mui/material/CircularProgress';

export default function LoadingWidget({ message }: { message: string }) {
  return (
    <div className="border-darkpink bg-lightpink fixed top-1/2 left-1/2 z-100 flex h-1/4 w-1/4 -translate-x-1/2 -translate-y-1/2 flex-col content-center items-center justify-center rounded-md border-3 p-2 tracking-wider text-black">
      <p>{message}</p>
      <CircularProgress sx={{ color: '#e1b3b5' }} />
    </div>
  );
}
