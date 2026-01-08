import { MediaType, PreSavedMediaItem } from 'lib/interfaces/globalInterfaces';
import { NextResponse } from 'next/server';

export interface ErrorResponse {
  error: string;
  message: string;
}

export interface SearchErrorResponse extends ErrorResponse {
  failedSearchData: { title: string; author: string } | [];
}

export interface DatabaseSaveEditErrorResponse extends ErrorResponse {
  errors: string[];
  title: string;
}

export class ApiError extends Error {
  status: number;
  error: string;
  constructor(status: number, error: string, message: string) {
    super(message);
    this.name = 'ApiError';
    this.status = status;
    this.error = error;
  }

  format(): NextResponse<ErrorResponse> {
    return NextResponse.json(
      { error: this.error, message: this.message },
      { status: this.status },
    );
  }
}

export class PGDatabaseError extends ApiError {
  actionAttemptItem: PreSavedMediaItem;
  type: MediaType;
  errorDetail: string;
  constructor(
    actionAttemptItem: PreSavedMediaItem,
    type: MediaType,
    errorDetail: string,
  ) {
    super(
      409,
      'Duplication Attempt Error',
      'You attempted to save an item to the database that already exists',
    );
    this.name = 'DuplicationAttemptError';
    this.errorDetail = errorDetail;
    this.actionAttemptItem = actionAttemptItem;
    this.type = type;
  }

  format(): NextResponse<DatabaseSaveEditErrorResponse> {
    return NextResponse.json(
      {
        error: this.error,
        message: this.message,
        errors: [this.errorDetail],
        actionAttemptItem: this.actionAttemptItem,
        type: this.type,
      },
      { status: this.status },
    );
  }
}

export class SchemaViolationError extends ApiError {
  schemaErrors: string[];
  actionAttemptItem: PreSavedMediaItem;
  type: MediaType;
  constructor(
    schemaErrors: string[],
    actionAttemptItem: PreSavedMediaItem,
    type: MediaType,
  ) {
    super(
      422,
      'Schema Violation',
      'Schema violation(s) during save/edit request',
    );
    this.name = 'SchemaViolationError';
    this.schemaErrors = this.formatErrors(schemaErrors);
    this.actionAttemptItem = actionAttemptItem;
    this.type = type;
  }

  formatErrors(schemaErrors: string[]): string[] {
    const missingFields: string[] = [];
    const wrongTypes: string[] = [];
    for (const error of schemaErrors) {
      if (error.includes('instance requires property')) {
        const missingField = error.split('"')[1];
        missingFields.push(`Save/Edit request missing ${missingField}`);
      } else if (error.includes('is not of a type(s)')) {
        const wrongTypeField = error.split(' ')[0].split('.')[1];
        wrongTypes.push(`${wrongTypeField} is of wrong type`);
      } else if (error.includes('does not meet minimum length')) {
        const field = error.split(' ')[0].split('.')[1];
        missingFields.push(`Save/Edit request missing ${field}`);
      }
    }
    return [...missingFields, ...wrongTypes];
  }

  format(): NextResponse<DatabaseSaveEditErrorResponse> {
    return NextResponse.json(
      {
        error: this.error,
        message: this.message,
        errors: this.schemaErrors,
        actionAttemptItem: this.actionAttemptItem,
        type: this.type,
      },
      { status: this.status },
    );
  }
}

export class OpenLibraryError extends ApiError {
  failedSearchData: { title: string; author: string };
  constructor(
    status: number,
    error: string,
    message: string,
    failedSearchData: { title: string; author: string },
  ) {
    super(status, error, message);
    this.failedSearchData = failedSearchData;
  }
  format(): NextResponse<ErrorResponse> {
    return NextResponse.json(
      {
        error: this.error,
        message: this.message,
        failedSearchData: this.failedSearchData,
      },
      { status: this.status },
    );
  }
}

export class GoogleSearchError extends ApiError {
  failedSearchData: [];
  constructor(status: number, error: string, message: string) {
    super(status, error, message);
    this.failedSearchData = [];
  }
  format(): NextResponse<ErrorResponse> {
    return NextResponse.json(
      {
        error: this.error,
        message: this.message,
        failedSearchData: this.failedSearchData,
      },
      { status: this.status },
    );
  }
}
