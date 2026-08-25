'use client';

import { useState } from 'react';

import Popover from '@mui/material/Popover';
import PopupState, { bindPopover, bindTrigger } from 'material-ui-popup-state';

import Button from '@/components/ui/button';
import {
  Command,
  CommandEmpty,
  CommandGroup,
  CommandInput,
  CommandItem,
  CommandList,
} from '@/components/ui/command';

export function Combobox({
  label,
  items,
  onSelect,
}: {
  label: string;
  items: { value: string }[];
  onSelect: (value: string) => void;
}) {
  const [search, setSearch] = useState('');

  return (
    <PopupState variant="popover" popupId="demo-popup-popover">
      {(popupState) => (
        <div>
          <Button
            {...bindTrigger(popupState)}
            width={200}
            fontSize={24}
            label={`Select ${label}...`}
            className="justify-between"
            variant="popover"
          />
          <Popover
            {...bindPopover(popupState)}
            anchorOrigin={{
              vertical: 'bottom',
              horizontal: 'center',
            }}
            transformOrigin={{
              vertical: 'top',
              horizontal: 'center',
            }}
          >
            <Command>
              <CommandInput
                placeholder={`Search ${label.toLowerCase()}...`}
                className="h-9"
                onValueChange={setSearch}
              />
              <CommandList>
                <CommandEmpty>No {label.toLowerCase()} found.</CommandEmpty>
                <CommandGroup>
                  {items
                    .filter((item) =>
                      item.value.toLowerCase().includes(search.toLowerCase()),
                    )
                    .map((item) => (
                      <CommandItem
                        key={item.value}
                        value={item.value}
                        onSelect={onSelect}
                      >
                        {item.value}
                      </CommandItem>
                    ))}
                </CommandGroup>
              </CommandList>
            </Command>
          </Popover>
        </div>
      )}
    </PopupState>
  );
}
