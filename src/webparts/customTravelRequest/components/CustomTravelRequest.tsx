import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ICustomTravelRequestProps } from './ICustomTravelRequestProps';
import styles from './CustomTravelRequest.module.scss';

interface IDetailLine {
  id: number;
  title: string;
  date: string;
  category: string;
  description: string;
  amount: number;
}

interface ITravelSummary {
  Id: number;
  Title: string;
  TravelID?: string;
  EstimatedCost?: number;
  TravelStatus?: string;
  DepartureDate?: string;
  ReturnDate?: string;
  Destination?: string;
  Created?: string;
}

interface ICustomTravelRequestState {
  currentView: 'dashboard' | 'form';
  requests: ITravelSummary[];
  loadingRequests: boolean;

  travelPurpose: string;
  travelType: string;
  destination: string;
  departureDate: string;
  returnDate: string;
  detailLines: IDetailLine[];
  
  // Approver selection
  resolvedManager: string | null;
  selectedApprover: string;
  approversList: string[];

  // New line item inputs
  newTitle: string;
  newDate: string;
  newCategory: string;
  newDescription: string;
  newAmount: string;

  // Metadata & Runtime States
  userId: number | null;
  travelTypes: string[];
  categories: string[];
  loadingConfig: boolean;
  isSubmitting: boolean;
  submitSuccess: boolean;
  submitError: string | null;
  createdTravelID: string | null;
}

export default class CustomTravelRequest extends React.Component<ICustomTravelRequestProps, ICustomTravelRequestState> {
  private _lineItemIdCounter = 0;

  constructor(props: ICustomTravelRequestProps) {
    super(props);

    // Default dates: departure tomorrow, return next week
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const returnDay = new Date();
    returnDay.setDate(returnDay.getDate() + 7);

    this.state = {
      currentView: 'dashboard',
      requests: [],
      loadingRequests: false,

      travelPurpose: '',
      travelType: '',
      destination: '',
      departureDate: tomorrow.toISOString().split('T')[0],
      returnDate: returnDay.toISOString().split('T')[0],
      detailLines: [],
      
      resolvedManager: null,
      selectedApprover: '',
      approversList: [],

      newTitle: '',
      newDate: tomorrow.toISOString().split('T')[0],
      newCategory: '',
      newDescription: '',
      newAmount: '',
      
      userId: null,
      travelTypes: ['Domestic', 'International'],
      categories: ['Flight', 'Hotel', 'Meals', 'Car Rental', 'Conference Fee', 'Train', 'Other'],
      loadingConfig: false,
      isSubmitting: false,
      submitSuccess: false,
      submitError: null,
      createdTravelID: null
    };
  }

  public componentDidMount(): void {
    this._resolveUserAndMetadata();
    this._fetchSubmittedRequests();
  }

  public componentDidUpdate(prevProps: ICustomTravelRequestProps, prevState: ICustomTravelRequestState): void {
    // If list configuration changed, refresh metadata and requests
    if (
      prevProps.properties.parentListId !== this.props.properties.parentListId ||
      prevProps.properties.childListId !== this.props.properties.childListId ||
      prevProps.properties.parentTravelTypeColumn !== this.props.properties.parentTravelTypeColumn ||
      prevProps.properties.childCategoryColumn !== this.props.properties.childCategoryColumn ||
      JSON.stringify(prevProps.properties.fallbackApprovers) !== JSON.stringify(this.props.properties.fallbackApprovers)
    ) {
      this._resolveUserAndMetadata();
      this._fetchSubmittedRequests();
    }

    // Refresh default selection when changed in properties
    if (prevProps.properties.defaultTravelType !== this.props.properties.defaultTravelType) {
      this.setState({ travelType: this.props.properties.defaultTravelType });
    }

    // Refresh requests when returning to the dashboard view
    if (prevState.currentView === 'form' && this.state.currentView === 'dashboard') {
      this._fetchSubmittedRequests();
    }
  }

  // Resolve SharePoint User ID, user manager profile, and choices
  private _resolveUserAndMetadata(): void {
    const { siteUrl, spHttpClient, properties } = this.props;

    if (!properties.parentListId || !properties.childListId) {
      return;
    }

    this.setState({ loadingConfig: true });

    // 1. Resolve current user ID
    const userUrl = `${siteUrl}/_api/web/currentUser`;
    const userPromise = spHttpClient.get(userUrl, SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => res.json())
      .then((data: { Id: number }) => data.Id)
      .catch(err => {
        console.error("Error resolving current user ID:", err);
        return null;
      });

    // 2. Fetch Parent Travel Type Choices
    let travelTypePromise = Promise.resolve(['Domestic', 'International']);
    if (properties.parentTravelTypeColumn) {
      const subUrl = `${siteUrl}/_api/web/lists(guid'${properties.parentListId}')/fields/getByInternalNameOrTitle('${properties.parentTravelTypeColumn}')`;
      travelTypePromise = spHttpClient.get(subUrl, SPHttpClient.configurations.v1)
        .then((res: SPHttpClientResponse) => res.json())
        .then((data: { Choices?: string[] }) => data.Choices || ['Domestic', 'International'])
        .catch(err => {
          console.warn("Failed to load parent choice fields, using default fallbacks:", err);
          return ['Domestic', 'International'];
        });
    }

    // 3. Fetch Child Category Choices
    let categoryPromise = Promise.resolve(['Flight', 'Hotel', 'Meals', 'Car Rental', 'Conference Fee', 'Train', 'Other']);
    if (properties.childCategoryColumn) {
      const catUrl = `${siteUrl}/_api/web/lists(guid'${properties.childListId}')/fields/getByInternalNameOrTitle('${properties.childCategoryColumn}')`;
      categoryPromise = spHttpClient.get(catUrl, SPHttpClient.configurations.v1)
        .then((res: SPHttpClientResponse) => res.json())
        .then((data: { Choices?: string[] }) => data.Choices || ['Flight', 'Hotel', 'Meals', 'Car Rental', 'Conference Fee', 'Train', 'Other'])
        .catch(err => {
          console.warn("Failed to load child choice fields, using default fallbacks:", err);
          return ['Flight', 'Hotel', 'Meals', 'Car Rental', 'Conference Fee', 'Train', 'Other'];
        });
    }

    // 4. Fetch manager from User Profile properties
    const profileUrl = `${siteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`;
    const managerPromise = spHttpClient.get(profileUrl, SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => res.json())
      .then((data: any) => {
        if (data && data.UserProfileProperties) {
          const managerProp = data.UserProfileProperties.find((p: any) => p.Key === 'Manager');
          let manager = managerProp ? managerProp.Value : null;
          if (manager) {
            // Manager value is claims string: "i:0#.f|membership|manager@tenant.onmicrosoft.com"
            if (manager.indexOf('|') > -1) {
              manager = manager.split('|').pop();
            }
            return manager;
          }
        }
        return null;
      })
      .catch(err => {
        console.warn("User Profile Service failed to retrieve manager:", err);
        return null;
      });

    Promise.all([userPromise, travelTypePromise, categoryPromise, managerPromise])
      .then(([userId, travelTypes, categories, resolvedManager]) => {
        // Parse fallback approvers from People Picker array
        let fallbackList = properties.fallbackApprovers || [];
        if (typeof fallbackList === 'string') {
          fallbackList = (fallbackList as string).split(',')
            .map(email => email.trim())
            .filter(email => email.length > 0)
            .map(email => ({ email, login: email, fullName: email } as any));
        }
        const approversList = (Array.isArray(fallbackList) ? fallbackList : [])
          .map(p => {
            let email = p.email || p.login || '';
            if (email.indexOf('|') > -1) {
              email = email.split('|').pop() || '';
            }
            return email.trim();
          })
          .filter(email => email.length > 0);

        // Pre-fill selection logic: resolved manager takes precedence
        const selectedApprover = resolvedManager || approversList[0] || '';

        this.setState({
          userId,
          travelTypes,
          categories,
          travelType: properties.defaultTravelType || travelTypes[0] || 'Domestic',
          newCategory: categories[0] || 'Flight',
          resolvedManager,
          selectedApprover,
          approversList,
          loadingConfig: false
        });
      })
      .catch(err => {
        console.error("Metadata resolution failed:", err);
        this.setState({ loadingConfig: false });
      });
  }

  // Fetch requests from SharePoint list
  private _fetchSubmittedRequests(): void {
    const { siteUrl, spHttpClient, properties, userEmail } = this.props;

    if (!properties.parentListId) {
      return;
    }

    this.setState({ loadingRequests: true });

    // Build select query for mapped columns
    const selectFields = ['Id', 'Title', 'Created'];
    const {
      parentTitleColumn,
      parentTravelIdColumn,
      parentEstimatedCostColumn,
      parentTravelStatusColumn,
      parentFromDateColumn,
      parentToDateColumn,
      parentTravelerEmailColumn
    } = properties;

    if (parentTitleColumn && parentTitleColumn !== 'Title') selectFields.push(parentTitleColumn);
    if (parentTravelIdColumn) selectFields.push(parentTravelIdColumn);
    if (parentEstimatedCostColumn) selectFields.push(parentEstimatedCostColumn);
    if (parentTravelStatusColumn) selectFields.push(parentTravelStatusColumn);
    if (parentFromDateColumn) selectFields.push(parentFromDateColumn);
    if (parentToDateColumn) selectFields.push(parentToDateColumn);
    if (parentTravelerEmailColumn) selectFields.push(parentTravelerEmailColumn);

    selectFields.push('Destination');

    const selectQuery = `$select=${selectFields.join(',')}`;
    
    // Filter requests submitted by current traveler if traveler email column is mapped
    let filterQuery = "";
    if (parentTravelerEmailColumn) {
      filterQuery = `&$filter=${parentTravelerEmailColumn} eq '${userEmail}'`;
    }

    const url = `${siteUrl}/_api/web/lists(guid'${properties.parentListId}')/items?${selectQuery}${filterQuery}&$orderby=Created desc&$top=50`;

    spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => res.json())
      .then((data: { value: any[] }) => {
        const requests = data.value.map(item => {
          const summary: ITravelSummary = {
            Id: item.Id,
            Title: item[parentTitleColumn || 'Title'] || item.Title,
            Created: item.Created,
            Destination: item.Destination || '—'
          };

          if (parentTravelIdColumn) summary.TravelID = item[parentTravelIdColumn];
          if (parentEstimatedCostColumn) summary.EstimatedCost = item[parentEstimatedCostColumn];
          if (parentTravelStatusColumn) {
            summary.TravelStatus = typeof item[parentTravelStatusColumn] === 'object' && item[parentTravelStatusColumn]
              ? item[parentTravelStatusColumn].Value
              : item[parentTravelStatusColumn];
          }
          if (parentFromDateColumn) summary.DepartureDate = item[parentFromDateColumn];
          if (parentToDateColumn) summary.ReturnDate = item[parentToDateColumn];

          return summary;
        });

        this.setState({ requests, loadingRequests: false });
      })
      .catch(err => {
        console.error("Error fetching submitted travel requests list:", err);
        this.setState({ loadingRequests: false });
      });
  }

  // Handle adding line item to details list in state
  private _handleAddLine = (e: React.FormEvent): void => {
    e.preventDefault();
    const { newTitle, newDate, newCategory, newDescription, newAmount, categories } = this.state;
    const { properties } = this.props;

    if (!newTitle.trim() || !newAmount || parseFloat(newAmount) <= 0) {
      alert("Please provide a valid Cost Item Title and a positive Amount.");
      return;
    }

    const newLine: IDetailLine = {
      id: ++this._lineItemIdCounter,
      title: newTitle,
      date: properties.showChildExpenseDate !== false ? newDate : new Date().toISOString().split('T')[0],
      category: properties.showChildCategory !== false ? newCategory : (categories[0] || 'Flight'),
      description: properties.showChildDescription !== false ? newDescription : '',
      amount: parseFloat(newAmount)
    };

    this.setState(prevState => ({
      detailLines: [...prevState.detailLines, newLine],
      newTitle: '',
      newAmount: '',
      newDescription: ''
    }));
  };

  // Remove line item from state
  private _handleRemoveLine = (id: number): void => {
    this.setState(prevState => ({
      detailLines: prevState.detailLines.filter(line => line.id !== id)
    }));
  };

  // Submit form data to SharePoint Lists
  private _handleSubmit = (): void => {
    const { travelPurpose, travelType, destination, departureDate, returnDate, detailLines, userId, travelTypes, selectedApprover } = this.state;
    const { siteUrl, spHttpClient, properties, userEmail } = this.props;

    // Default title if purpose field is hidden
    const finalPurpose = properties.showParentTitle !== false && travelPurpose.trim() 
      ? travelPurpose 
      : `Travel Request - ${destination || 'Trip'} - ${departureDate}`;

    if (!destination.trim()) {
      alert("Please specify a Destination.");
      return;
    }

    if (new Date(departureDate) > new Date(returnDate)) {
      alert("Departure Date cannot be after Return Date.");
      return;
    }

    if (!selectedApprover) {
      alert("Please configure or select an Approver.");
      return;
    }

    if (detailLines.length === 0) {
      alert("Please add at least one estimated cost item.");
      return;
    }

    this.setState({ isSubmitting: true, submitError: null });

    const totalEstimatedCost = detailLines.reduce((sum, item) => sum + item.amount, 0);

    // Stage 1: Post to Parent List
    const parentUrl = `${siteUrl}/_api/web/lists(guid'${properties.parentListId}')/items`;
    const parentPayload: any = {};
    
    parentPayload[properties.parentTitleColumn || 'Title'] = finalPurpose;
    parentPayload['Destination'] = destination;
    
    if (properties.parentTravelTypeColumn) {
      parentPayload[properties.parentTravelTypeColumn] = travelType;
    }
    if (properties.parentEstimatedCostColumn) {
      parentPayload[properties.parentEstimatedCostColumn] = totalEstimatedCost;
    }
    if (properties.parentTravelerEmailColumn) {
      parentPayload[properties.parentTravelerEmailColumn] = userEmail;
    }
    if (properties.parentFromDateColumn) {
      parentPayload[properties.parentFromDateColumn] = new Date(departureDate).toISOString();
    }
    if (properties.parentToDateColumn) {
      parentPayload[properties.parentToDateColumn] = new Date(returnDate).toISOString();
    }
    if (properties.parentTravelerNameColumn && userId !== null) {
      parentPayload[`${properties.parentTravelerNameColumn}Id`] = userId;
    }
    if (properties.parentTravelStatusColumn) {
      parentPayload[properties.parentTravelStatusColumn] = properties.defaultTravelStatus || 'Submitted';
    }
    // Approver is written directly to the standard 'Approver' text field
    parentPayload['Approver'] = selectedApprover;


    const postHeaders = {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=nometadata',
      'odata-version': ''
    };

    spHttpClient.post(parentUrl, SPHttpClient.configurations.v1, {
      headers: postHeaders,
      body: JSON.stringify(parentPayload)
    })
      .then((res: SPHttpClientResponse) => {
        if (!res.ok) {
          return res.json().then(err => { throw err; });
        }
        return res.json();
      })
      .then((parentItem: { Id: number }) => {
        const parentId = parentItem.Id;

        // Stage 2: Post Child lines to details list
        const childUrl = `${siteUrl}/_api/web/lists(guid'${properties.childListId}')/items`;
        const promises = detailLines.map(line => {
          const childPayload: any = {};
          
          childPayload[properties.childTitleColumn || 'Title'] = line.title;
          
          if (properties.childExpenseDateColumn) {
            childPayload[properties.childExpenseDateColumn] = new Date(line.date).toISOString();
          }
          if (properties.childCategoryColumn) {
            childPayload[properties.childCategoryColumn] = line.category;
          }
          if (properties.childAmountColumn) {
            childPayload[properties.childAmountColumn] = line.amount;
          }
          if (properties.childDescriptionColumn && line.description) {
            childPayload[properties.childDescriptionColumn] = line.description;
          }
          if (properties.childParentIdColumn) {
            // Write to Lookup Field using ID
            childPayload[`${properties.childParentIdColumn}Id`] = parentId;
          }

          return spHttpClient.post(childUrl, SPHttpClient.configurations.v1, {
            headers: postHeaders,
            body: JSON.stringify(childPayload)
          }).then((childRes: SPHttpClientResponse) => {
            if (!childRes.ok) {
              return childRes.json().then(err => { throw err; });
            }
            return childRes.json();
          });
        });

        return Promise.all(promises).then(() => parentId);
      })
      .then((parentId) => {
        this.setState({
          isSubmitting: false,
          submitSuccess: true,
          createdTravelID: `TRV-${new Date().getFullYear()}-${parentId}`,
          travelPurpose: '',
          destination: '',
          detailLines: []
        });
      })
      .catch(err => {
        console.error("Error submitting travel request transaction:", err);
        const errMsg = err.error && err.error.message ? err.error.message.value : (err.message || "An unexpected error occurred during submission.");
        this.setState({
          isSubmitting: false,
          submitError: errMsg
        });
      });
  };

  private _closeSuccessScreen = (): void => {
    this.setState({
      submitSuccess: false,
      createdTravelID: null,
      submitError: null,
      currentView: 'dashboard'
    });
  };

  private _goToFormView = (): void => {
    this.setState({
      currentView: 'form',
      submitSuccess: false,
      submitError: null
    });
  };

  private _goToDashboardView = (): void => {
    this.setState({
      currentView: 'dashboard',
      submitSuccess: false,
      submitError: null
    });
  };

  public render(): React.ReactElement<ICustomTravelRequestProps> {
    const { properties, userEmail, userName } = this.props;
    const {
      currentView,
      requests,
      loadingRequests,
      travelPurpose,
      travelType,
      destination,
      departureDate,
      returnDate,
      detailLines,
      resolvedManager,
      selectedApprover,
      approversList,
      newTitle,
      newDate,
      newCategory,
      newDescription,
      newAmount,
      travelTypes,
      categories,
      loadingConfig,
      isSubmitting,
      submitSuccess,
      submitError,
      createdTravelID
    } = this.state;

    // Guard: Webpart properties configuration check
    if (!properties.parentListId || !properties.childListId) {
      return (
        <div className={styles.container}>
          <div className={styles.configPlaceholder}>
            <div className={styles.placeholderIcon}>✈️</div>
            <h2 className={styles.placeholderTitle}>Setup Required</h2>
            <p className={styles.placeholderText}>
              Please configure the Parent Travel List and Child Details List column mappings inside the Web Part Property Pane to enable the transaction form.
            </p>
          </div>
        </div>
      );
    }

    if (loadingConfig) {
      return (
        <div className={styles.container}>
          <div className={styles.loadingWrapper}>
            <div className={styles.spinner}></div>
            <p>Fetching list schemas and traveler profile...</p>
          </div>
        </div>
      );
    }

    // SUCCESS SUBMISSION SCREEN
    if (submitSuccess) {
      return (
        <div className={styles.container}>
          <div className={styles.successWrapper}>
            <div className={styles.successIcon}>✓</div>
            <h2 className={styles.successTitle}>Travel Request Submitted!</h2>
            <p className={styles.successText}>
              Your travel booking and cost request has been registered.
            </p>
            <div className={styles.receiptCard}>
              <div className={styles.receiptRow}>
                <span className={styles.receiptLabel}>Assigned ID:</span>
                <span className={styles.receiptValue}>{createdTravelID || "Generating..."}</span>
              </div>
              <div className={styles.receiptRow}>
                <span className={styles.receiptLabel}>Approver:</span>
                <span className={styles.receiptValue}>{selectedApprover}</span>
              </div>
              <div className={styles.receiptRow}>
                <span className={styles.receiptLabel}>Initial Status:</span>
                <span className={`${styles.receiptValue} ${styles.statusBadge} ${styles.badgeSubmitted}`}>Submitted</span>
              </div>
              <p className={styles.noticeText}>
                The Power Automate approval cycle is active. Your manager/approver has been notified to review and approve/reject this request.
              </p>
            </div>
            <button className={styles.primaryButton} onClick={this._closeSuccessScreen}>
              Back to Dashboard
            </button>
          </div>
        </div>
      );
    }

    // DASHBOARD VIEW
    if (currentView === 'dashboard') {
      const activeTripsCount = requests.filter(r => r.TravelStatus !== 'Rejected' && r.TravelStatus !== 'Completed').length;
      const totalBudget = requests.reduce((sum, r) => sum + (r.EstimatedCost || 0), 0);
      const pendingCount = requests.filter(r => r.TravelStatus === 'Submitted' || r.TravelStatus === 'Pending Approval').length;
      
      return (
        <div className={styles.container}>
          <div className={styles.formHeader}>
            <div className={styles.headerFlex}>
              <div>
                <h1 className={styles.formTitle}>
                  <span>✈️</span> Corporate Travel Request Portal
                </h1>
                <p className={styles.formSubtitle}>Submit new travel plans or review active approval statuses</p>
              </div>
              <button className={styles.newClaimBtn} onClick={this._goToFormView}>
                + Plan New Journey
              </button>
            </div>
          </div>

          {/* Stats Bar */}
          <div className={styles.statsRow}>
            <div className={styles.statCard}>
              <div className={styles.statIcon}>🗺️</div>
              <div className={styles.statMeta}>
                <span className={styles.statVal}>{activeTripsCount}</span>
                <span className={styles.statLabel}>Active Requests</span>
              </div>
            </div>
            <div className={styles.statCard}>
              <div className={`${styles.statIcon} cost`}>💵</div>
              <div className={styles.statMeta}>
                <span className={styles.statVal}>${totalBudget.toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}</span>
                <span className={styles.statLabel}>Total Estimated Budget</span>
              </div>
            </div>
            <div className={styles.statCard}>
              <div className={`${styles.statIcon} pending`}>⏳</div>
              <div className={styles.statMeta}>
                <span className={styles.statVal}>{pendingCount}</span>
                <span className={styles.statLabel}>Awaiting Decision</span>
              </div>
            </div>
          </div>

          <div className={styles.card}>
            <h2 className={styles.cardTitle}>
              <span>📋</span> My Travel Log & Statuses
            </h2>
            
            {loadingRequests ? (
              <div className={styles.loadingWrapper}>
                <div className={styles.spinner}></div>
                <p>Loading your travel history...</p>
              </div>
            ) : requests.length === 0 ? (
              <div className={styles.emptyDashboard}>
                <div className={styles.emptyIcon}>🌍</div>
                <h3>No trips submitted yet</h3>
                <p>Ready to go? Click "Plan New Journey" to log your travel details.</p>
              </div>
            ) : (
              <div className={styles.tableWrapper}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th>Travel ID</th>
                      <th>Purpose / Title</th>
                      <th>Destination</th>
                      <th>Dates</th>
                      <th style={{ textAlign: 'right' }}>Est. Cost</th>
                      <th style={{ textAlign: 'center' }}>Approval Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {requests.map(req => {
                      const depDate = req.DepartureDate ? new Date(req.DepartureDate).toLocaleDateString() : '—';
                      const retDate = req.ReturnDate ? new Date(req.ReturnDate).toLocaleDateString() : '—';
                      const status = req.TravelStatus || 'Submitted';
                      
                      let statusClass = styles.badgeSubmitted;
                      if (status === 'Approved') statusClass = styles.badgeApproved;
                      else if (status === 'Rejected') statusClass = styles.badgeRejected;
                      else if (status === 'Pending Approval') statusClass = styles.badgePending;
                      else if (status === 'Completed') statusClass = styles.badgeCompleted;

                      return (
                        <tr key={req.Id}>
                          <td style={{ fontWeight: 700, color: '#1e3a8a' }}>
                            {req.TravelID || "Pending ID..."}
                          </td>
                          <td style={{ fontWeight: 600 }}>{req.Title}</td>
                          <td>{req.Destination}</td>
                          <td style={{ color: '#64748b' }}>
                            {depDate} – {retDate}
                          </td>
                          <td style={{ textAlign: 'right', fontWeight: 700, color: '#0f172a' }}>
                            {req.EstimatedCost !== undefined ? `$${req.EstimatedCost.toFixed(2)}` : '—'}
                          </td>
                          <td style={{ textAlign: 'center' }}>
                            <span className={`${styles.statusBadge} ${statusClass}`}>
                              {status}
                            </span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      );
    }

    // FORM VIEW
    const totalRequestCost = detailLines.reduce((sum, item) => sum + item.amount, 0);

    return (
      <div className={styles.container}>
        <div className={styles.formHeader}>
          <div className={styles.headerFlex}>
            <div>
              <h1 className={styles.formTitle}>
                <span>✈️</span> New Travel Request Ingestion
              </h1>
              <p className={styles.formSubtitle}>Log your flight, hotel, and details to request manager verification</p>
            </div>
            <button className={styles.backBtn} onClick={this._goToDashboardView} disabled={isSubmitting}>
              ← Back to Log
            </button>
          </div>
        </div>

        {submitError && (
          <div className={styles.errorBanner}>
            <strong>Submission Error:</strong> {submitError}
          </div>
        )}

        <div className={styles.formLayout}>
          {/* Section 1: Parent Travel details */}
          <div className={styles.card}>
            <h2 className={styles.cardTitle}>
              <span>ℹ️</span> 1. General Travel Details
            </h2>
            
            <div className={styles.formFieldsGrid}>
              <div className={styles.inputGroup}>
                <label className={styles.label}>Traveler Name</label>
                <input type="text" className={styles.inputText} value={userName} disabled={true} />
              </div>

              <div className={styles.inputGroup}>
                <label className={styles.label}>Traveler Email</label>
                <input type="text" className={styles.inputText} value={userEmail} disabled={true} />
              </div>

              {/* Approver Pre-fill / Fallback Dropdown */}
              {resolvedManager ? (
                <div className={styles.inputGroup}>
                  <label className={styles.label}>Approver (Resolved Manager) *</label>
                  <input
                    type="text"
                    className={styles.inputText}
                    value={selectedApprover}
                    disabled={true}
                  />
                </div>
              ) : (
                <div className={styles.inputGroup}>
                  <label className={styles.label}>Approver (Select Fallback) *</label>
                  <select
                    className={styles.inputSelect}
                    value={selectedApprover}
                    onChange={e => this.setState({ selectedApprover: e.target.value })}
                    disabled={isSubmitting}
                  >
                    {approversList.map(email => (
                      <option key={email} value={email}>{email}</option>
                    ))}
                    {approversList.length === 0 && (
                      <option value="">-- No Fallback Approvers Configured --</option>
                    )}
                  </select>
                </div>
              )}

              {properties.showParentTitle !== false && (
                <div className={styles.inputGroup}>
                  <label className={styles.label}>Trip Purpose *</label>
                  <input
                    type="text"
                    className={styles.inputText}
                    placeholder="e.g., Q3 Global Sales Summit"
                    value={travelPurpose}
                    onChange={e => this.setState({ travelPurpose: e.target.value })}
                    disabled={isSubmitting}
                  />
                </div>
              )}

              <div className={styles.inputGroup}>
                <label className={styles.label}>Destination City/Country *</label>
                <input
                  type="text"
                  className={styles.inputText}
                  placeholder="e.g. London, UK"
                  value={destination}
                  onChange={e => this.setState({ destination: e.target.value })}
                  disabled={isSubmitting}
                />
              </div>

              {properties.parentTravelTypeColumn && (
                <div className={styles.inputGroup}>
                  <label className={styles.label}>Travel Type *</label>
                  <select
                    className={styles.inputSelect}
                    value={travelType}
                    onChange={e => this.setState({ travelType: e.target.value })}
                    disabled={isSubmitting}
                  >
                    {travelTypes.map(t => (
                      <option key={t} value={t}>{t}</option>
                    ))}
                  </select>
                </div>
              )}

              <div className={styles.inputGroup}>
                <label className={styles.label}>Departure Date *</label>
                <input
                  type="date"
                  className={styles.inputDate}
                  value={departureDate}
                  onChange={e => this.setState({ departureDate: e.target.value })}
                  disabled={isSubmitting}
                />
              </div>

              <div className={styles.inputGroup}>
                <label className={styles.label}>Return Date *</label>
                <input
                  type="date"
                  className={styles.inputDate}
                  value={returnDate}
                  onChange={e => this.setState({ returnDate: e.target.value })}
                  disabled={isSubmitting}
                />
              </div>
            </div>
          </div>

          {/* Section 2: Inline grid inputs */}
          <div className={styles.card}>
            <h2 className={styles.cardTitle}>
              <span>💰</span> 2. Travel Expense Cost Estimation
            </h2>
            
            <form onSubmit={this._handleAddLine} className={styles.gridForm}>
              <div className={styles.gridInputs} style={{ 
                gridTemplateColumns: `2fr ${properties.showChildExpenseDate !== false ? '1.2fr' : ''} ${properties.showChildCategory !== false ? '1.2fr' : ''} 1fr`.replace(/\s+/g, ' ').trim() 
              }}>
                <div className={styles.inputCol}>
                  <label className={styles.label}>Cost Item Title *</label>
                  <input
                    type="text"
                    className={styles.inputText}
                    placeholder="e.g. Flight ticket booking"
                    value={newTitle}
                    onChange={e => this.setState({ newTitle: e.target.value })}
                    disabled={isSubmitting}
                  />
                </div>
                
                {properties.showChildExpenseDate !== false && (
                  <div className={styles.inputCol}>
                    <label className={styles.label}>Date *</label>
                    <input
                      type="date"
                      className={styles.inputDate}
                      value={newDate}
                      onChange={e => this.setState({ newDate: e.target.value })}
                      disabled={isSubmitting}
                    />
                  </div>
                )}
                
                {properties.showChildCategory !== false && (
                  <div className={styles.inputCol}>
                    <label className={styles.label}>Cost Type Category *</label>
                    <select
                      className={styles.inputSelect}
                      value={newCategory}
                      onChange={e => this.setState({ newCategory: e.target.value })}
                      disabled={isSubmitting}
                    >
                      {categories.map(c => (
                        <option key={c} value={c}>{c}</option>
                      ))}
                    </select>
                  </div>
                )}
                
                <div className={styles.inputCol}>
                  <label className={styles.label}>Estimated Cost ($) *</label>
                  <input
                    type="number"
                    step="0.01"
                    className={styles.inputText}
                    placeholder="0.00"
                    value={newAmount}
                    onChange={e => this.setState({ newAmount: e.target.value })}
                    disabled={isSubmitting}
                  />
                </div>
              </div>
              
              <div className={styles.inputRow}>
                {properties.showChildDescription !== false ? (
                  <div className={styles.descCol}>
                    <label className={styles.label}>Detailed Description</label>
                    <textarea
                      className={styles.inputTextArea}
                      placeholder="Add travel flight codes, hotel name or booking remarks..."
                      value={newDescription}
                      onChange={e => this.setState({ newDescription: e.target.value })}
                      rows={2}
                      disabled={isSubmitting}
                    />
                  </div>
                ) : <div style={{ flex: 1 }}></div>}
                
                <div className={styles.buttonCol}>
                  <button type="submit" className={styles.secondaryButton} disabled={isSubmitting}>
                    + Add Cost Line
                  </button>
                </div>
              </div>
            </form>

            {/* List Table */}
            <div className={styles.tableWrapper}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    {properties.showChildExpenseDate !== false && <th>Date</th>}
                    <th>Cost Title</th>
                    {properties.showChildCategory !== false && <th>Category</th>}
                    {properties.showChildDescription !== false && <th>Description Details</th>}
                    <th style={{ textAlign: 'right' }}>Amount</th>
                    <th style={{ width: '50px' }}></th>
                  </tr>
                </thead>
                <tbody>
                  {detailLines.length === 0 ? (
                    <tr>
                      <td colSpan={6} className={styles.emptyRow}>
                        No estimated costs added yet. Use the fields above to add flights, lodging, etc.
                      </td>
                    </tr>
                  ) : (
                    detailLines.map(line => (
                      <tr key={line.id}>
                        {properties.showChildExpenseDate !== false && <td>{line.date}</td>}
                        <td style={{ fontWeight: 600 }}>{line.title}</td>
                        {properties.showChildCategory !== false && (
                          <td>
                            <span className={styles.pillBadge}>{line.category}</span>
                          </td>
                        )}
                        {properties.showChildDescription !== false && (
                          <td className={styles.descCell}>{line.description || "—"}</td>
                        )}
                        <td style={{ textAlign: 'right', fontWeight: 700, color: '#0f172a' }}>
                          ${line.amount.toFixed(2)}
                        </td>
                        <td style={{ textAlign: 'center' }}>
                          <button
                            type="button"
                            className={styles.trashBtn}
                            onClick={() => this._handleRemoveLine(line.id)}
                            disabled={isSubmitting}
                          >
                            🗑️
                          </button>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            <div className={styles.summaryContainer}>
              <div className={styles.summaryValueWrapper}>
                <span className={styles.summaryLabel}>Total Estimated Request Budget</span>
                <span className={styles.summaryValue}>${totalRequestCost.toFixed(2)}</span>
              </div>
              
              <button
                type="button"
                className={styles.primaryButton}
                onClick={this._handleSubmit}
                disabled={isSubmitting || detailLines.length === 0}
              >
                {isSubmitting ? (
                  <div className={styles.btnSpinnerWrapper}>
                    <div className={styles.btnSpinner}></div>
                    <span>Submitting Request...</span>
                  </div>
                ) : (
                  <span>Submit Travel Request</span>
                )}
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
