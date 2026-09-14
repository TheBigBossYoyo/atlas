import React from 'react'
import { AlertTriangle } from 'lucide-react'

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
        // UX-17 — themed, centered fallback (was an unstyled div/pre/button).
        <div className="viewer-fallback viewer-error" role="alert">
          <AlertTriangle size={32} className="viewer-fallback__icon" aria-hidden="true" />
          <h2 className="viewer-fallback__title">Viewer crashed</h2>
          <p className="viewer-fallback__detail">{error.message}</p>
          <button className="viewer-fallback__btn" onClick={this.handleReset}>
            Try again
          </button>
        </div>
      )
    }
    return this.props.children
  }
}
