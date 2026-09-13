import React from 'react'

type ViewerErrorBoundaryProps = {
  readonly children: React.ReactNode
  readonly onReset?: () => void
}

type ViewerErrorBoundaryState = {
  error: Error | null
}

export class ViewerErrorBoundary extends React.Component<
  ViewerErrorBoundaryProps,
  ViewerErrorBoundaryState
> {
  constructor(props: ViewerErrorBoundaryProps) {
    super(props)
    this.state = { error: null }
    this.handleReset = this.handleReset.bind(this)
  }

  static getDerivedStateFromError(error: Error): ViewerErrorBoundaryState {
    return { error }
  }

  componentDidCatch(error: Error, info: React.ErrorInfo): void {
    this.setState({ error })
    console.error('ViewerErrorBoundary caught:', error, info)
  }

  handleReset(): void {
    this.setState({ error: null })
    this.props.onReset?.()
  }

  render(): React.ReactNode {
    const { error } = this.state
    if (error !== null) {
      return (
        <div className="viewer-error" role="alert">
          <h2>Viewer crashed</h2>
          <pre>{error.message}</pre>
          <button onClick={this.handleReset}>Try again</button>
        </div>
      )
    }
    return this.props.children
  }
}
