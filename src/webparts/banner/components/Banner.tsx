import * as React from 'react';
import styles from './Banner.module.scss';
import { IBannerProps } from './IBannerProps';
import { Text } from 'office-ui-fabric-react';

export default class Banner extends React.Component<IBannerProps, {}> {
  public render(): React.ReactElement<IBannerProps> {
    const {
    } = this.props;

    return (
      <section className={`${styles.banner}`}>
        <Text>Hello Banner</Text>
      </section>
    );
  }
}
