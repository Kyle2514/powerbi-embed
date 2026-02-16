import * as React from 'react';
import { service, IEmbedConfiguration, models, factories } from 'powerbi-client';

interface IPowerBiEmbedProps {
  userEmail: string;
  reportId: string;
  datasetId: string;
  hasRLS: boolean;
  reportWorkspaceId: string;
  datasetWorkspaceId: string;
}

interface IPowerBIErrorDetail {
  message: string;
  detail?: string;
}

const powerbiService = new service.Service(
  factories.hpmFactory,
  factories.wpmpFactory,
  factories.routerFactory
);

const PowerBIEmbed: React.FC<IPowerBiEmbedProps> = ({ userEmail, reportId, datasetId, hasRLS, reportWorkspaceId, datasetWorkspaceId, }) => {
  const reportRef = React.useRef<HTMLDivElement>(null);
  const [error, setError] = React.useState<string | null>(null);

  React.useEffect(() => {
    if (!userEmail || !reportId || !datasetId || !reportWorkspaceId || !datasetWorkspaceId) {
      setError('report ID, dataset ID, report Workspace Id, and dataset Workspace Id are required to embed the report.');
      return;
    }

    if (reportRef.current) {
      powerbiService.reset(reportRef.current);
    }

    const fetchAndEmbedReport = async (): Promise<void> => {
      try {
        const response = await fetch(
          //`https://powerbi-embedded-digilab-fggbfbaeb0f5dvhx.francecentral-01.azurewebsites.net/api/embed-info?userEmail=${encodeURIComponent(userEmail)}&reportId=${encodeURIComponent(reportId)}&datasetId=${encodeURIComponent(datasetId)}&hasRLS=${hasRLS ? 'yes' : 'no'}`
          `http://localhost:3000/api/embed-info?userEmail=${encodeURIComponent(userEmail)}&reportId=${encodeURIComponent(reportId)}&datasetId=${encodeURIComponent(datasetId)}&hasRLS=${hasRLS ? 'yes' : 'no'}&reportWorkspaceId=${encodeURIComponent(reportWorkspaceId)}&datasetWorkspaceId=${encodeURIComponent(datasetWorkspaceId)}`
        );
        if (!response.ok) {
          let errorMessage = `HTTP error! status: ${response.status}`;
          try {
            const errorData = await response.json();
            errorMessage = errorData?.error?.message || errorData?.error || errorMessage;
          } catch {
            // ignore JSON parse error
          }
          throw new Error(errorMessage);
        }

        const embedInfo = await response.json();

        if (reportRef.current) {
          const embedConfig: IEmbedConfiguration = {
            type: 'report',
            id: embedInfo.reportId,
            embedUrl: embedInfo.embedUrl,
            accessToken: embedInfo.embedToken,
            tokenType: models.TokenType.Embed,
            settings: {
              panes: {
                filters: { visible: false },
                pageNavigation: { visible: false },
              },
              background: models.BackgroundType.Transparent,
            },
          };

          const report = powerbiService.embed(reportRef.current, embedConfig);

          report.on('loaded', () => {
            console.log('Power BI Report loaded');
          });

          report.on('rendered', () => {
            console.log('Power BI Report rendered');
          });

          report.on('error', (event: service.ICustomEvent<IPowerBIErrorDetail>) => {
            console.error('Power BI Embedding Error:', event.detail);
            setError(`Power BI Error: ${event.detail.message || JSON.stringify(event.detail)}`);
          });
        }
      } catch (err) {
        const errorMessage = err instanceof Error ? err.message : JSON.stringify(err);
        console.error('Error fetching embed info or embedding report:', errorMessage);
        setError(`Failed to load report: ${errorMessage}`);
      }
    };

    fetchAndEmbedReport().catch(err => {
      console.error('Error in fetchAndEmbedReport:', err);
    });

    return () => {
      if (reportRef.current) {
        powerbiService.reset(reportRef.current);
      }
    };
  }, [userEmail, reportId, datasetId, hasRLS, reportWorkspaceId, datasetWorkspaceId]);

  return (
    <div>
      {error && <div style={{ color: 'red', padding: '10px' }}>{error}</div>}
      <div ref={reportRef} style={{ height: '600px', width: '100%' }} />
    </div>
  );
};

export default PowerBIEmbed;
