import { Command as CommandPrimitive } from 'cmdk';
import SearchIcon from '@mui/icons-material/Search';

import { cn } from '@/lib/utils/classnames';

const Command = ({
  className,
  ...props
}: React.ComponentProps<typeof CommandPrimitive>) => (
  <CommandPrimitive
    className={cn(
      'bg-background text-foreground flex size-full flex-col overflow-hidden rounded-md',
      className,
    )}
    {...props}
  />
);

const CommandInput = ({
  className,
  ...props
}: React.ComponentProps<typeof CommandPrimitive.Input>) => (
  <div
    className={cn('flex items-center rounded-t-md border-b px-3', className)}
    cmdk-input-wrapper=""
  >
    <SearchIcon className="mr-2 size-4 shrink-0 opacity-50" />
    <CommandPrimitive.Input
      className={cn(
        'placeholder:text-foreground-weakest flex w-full bg-transparent py-2 outline-hidden disabled:cursor-not-allowed disabled:opacity-50 max-sm:text-[16px]',
      )}
      {...props}
    />
  </div>
);

const CommandList = ({
  className,
  ...props
}: React.ComponentProps<typeof CommandPrimitive.List>) => (
  <CommandPrimitive.List
    className={cn(
      'max-h-[300px] overflow-x-hidden overflow-y-auto [scroll-padding-block-end:8px] [scroll-padding-block-start:8px]',
      className,
    )}
    {...props}
  />
);

const CommandEmpty = ({
  className,
  ...props
}: React.ComponentProps<typeof CommandPrimitive.Empty>) => (
  <CommandPrimitive.Empty
    className={cn('text-foreground-weaker p-3 text-sm', className)}
    {...props}
  />
);

const CommandGroup = ({
  className,
  ...props
}: React.ComponentProps<typeof CommandPrimitive.Group>) => (
  <CommandPrimitive.Group
    className={cn(
      'text-foreground [&_[cmdk-group-heading]]:bg-darken-weaker [&_[cmdk-group-heading]]:text-foreground-weak overflow-hidden [&_[cmdk-group-heading]]:px-2 [&_[cmdk-group-heading]]:py-1.5 [&_[cmdk-group-heading]]:text-sm [&_[cmdk-group-heading]]:font-medium [&_[cmdk-group-heading]]:tracking-wider [&_[cmdk-group-heading]]:uppercase [&:not(:first-child)_[cmdk-group-heading]]:mt-1.5',
      className,
    )}
    {...props}
  />
);

const CommandItem = ({
  className,
  ...props
}: React.ComponentProps<typeof CommandPrimitive.Item>) => (
  <CommandPrimitive.Item
    className={cn(
      'data-[selected=true]:bg-accent data-[selected=true]:text-accent-foreground relative flex cursor-pointer items-center gap-2 rounded-xs px-2 py-1.5 outline-hidden select-none data-[disabled=true]:pointer-events-none data-[disabled=true]:opacity-50 max-sm:text-[16px] [&_svg]:pointer-events-none [&_svg]:size-4 [&_svg]:shrink-0',
      className,
    )}
    {...props}
  />
);

export {
  Command,
  CommandEmpty,
  CommandGroup,
  CommandInput,
  CommandItem,
  CommandList,
};
